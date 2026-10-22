"""Airfare alert: tracks economy NON-STOP fares via SerpAPI Google Flights and
notifies a Teams channel (Power Automate webhook) when the price drops or rises.

Trips are defined in trips.json (return / oneway / multicity). Usage:
  python airfare_alert.py check          # run on a schedule
  python airfare_alert.py list
  python airfare_alert.py test           # send a sample card to Teams (no search credit used)
  python airfare_alert.py add --type return --from SIN --to BOM --depart 2026-12-01 --return 2026-12-08
  python airfare_alert.py add --type oneway --from SIN --to BOM --depart 2026-12-01
  python airfare_alert.py add --type multicity --legs SIN-BOM:2026-12-01 BOM-DEL:2026-12-05 DEL-SIN:2026-12-09
  python airfare_alert.py remove <id>
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import requests

BASE = Path(__file__).parent
TRIPS_FILE = BASE / "trips.json"
STATE_FILE = BASE / "state.json"
SYMBOLS = {"INR": "₹", "SGD": "S$", "USD": "$", "EUR": "€", "GBP": "£"}
MONTHLY_CAP = int(os.getenv("MONTHLY_SEARCH_CAP", "240"))  # SerpAPI plan: 250/month, keep a buffer
NIGHT_START, NIGHT_END = 0, 4  # 00:00-04:00 local (Singapore) time: checked hourly

# ----------------------------------------------------------------- helpers
def load_env():
    env = BASE / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

def read(path, default):
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else default

def write(path, data):
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")

def money(v, cur):
    return f"{SYMBOLS.get(cur, cur + ' ')}{v:,.0f}"

def nice_time(s):
    """'2026-11-24 08:45' -> 'Tue 24 Nov, 08:45'"""
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M").strftime("%a %d %b, %H:%M")
    except ValueError:
        return s

def nice_date(s):
    return datetime.strptime(s, "%Y-%m-%d").strftime("%d %b %Y")

def hm(minutes):
    return f"{minutes // 60}h {minutes % 60:02d}m"

def budget_ok(state):
    """Count SerpAPI searches per calendar month; refuse once the cap is reached."""
    month = time.strftime("%Y-%m")
    u = state.get("_usage", {})
    if u.get("month") != month:
        u = {"month": month, "count": 0}
    state["_usage"] = u
    return u["count"] < MONTHLY_CAP

def legs_of(trip):
    """Normalise a trip into a list of (origin, destination, date) legs."""
    if trip["type"] == "oneway":
        return [(trip["from"], trip["to"], trip["depart"])]
    if trip["type"] == "return":
        return [(trip["from"], trip["to"], trip["depart"]), (trip["to"], trip["from"], trip["return"])]
    return [tuple(l) for l in trip["legs"]]

def describe(trip):
    return " | ".join(f"{o}->{d} {dt}" for o, d, dt in legs_of(trip)) + f" ({trip['type']}, economy, non-stop)"

def load_trips():
    """Read trips.json; trips without an "id" get a stable one derived from their route/dates."""
    trips = read(TRIPS_FILE, [])
    for t in trips:
        t.setdefault("id", "-".join(f"{o}{d}{dt}" for o, d, dt in legs_of(t)))
    return trips

# ----------------------------------------------------------------- search
def search(trip, extra=None):
    params = {
        "engine": "google_flights",
        "api_key": os.environ["SERPAPI_KEY"],
        "currency": os.getenv("CURRENCY", "INR"),
        "travel_class": 1,  # economy
        "stops": 1,         # non-stop only
        "hl": "en",
    }
    legs = legs_of(trip)
    if trip["type"] == "multicity":
        params["type"] = 3
        params["multi_city_json"] = json.dumps(
            [{"departure_id": o, "arrival_id": d, "date": dt} for o, d, dt in legs]
        )
    else:
        params.update(departure_id=legs[0][0], arrival_id=legs[0][1], outbound_date=legs[0][2])
        if trip["type"] == "return":
            params.update(type=1, return_date=trip["return"])
        else:
            params["type"] = 2
    params.update(extra or {})
    r = requests.get("https://serpapi.com/search.json", params=params, timeout=90)
    r.raise_for_status()
    data = r.json()
    if data.get("error"):
        raise RuntimeError(data["error"])
    return data

def summarise(flight):
    segs = flight["flights"]
    return {
        "price": float(flight["price"]),
        "airline": " + ".join(dict.fromkeys(s["airline"] for s in segs)),
        "numbers": ", ".join(s.get("flight_number", "") for s in segs),
        "dep": nice_time(segs[0]["departure_airport"]["time"]),
        "arr": nice_time(segs[-1]["arrival_airport"]["time"]),
        "from": segs[0]["departure_airport"]["id"],
        "to": segs[-1]["arrival_airport"]["id"],
        "duration": hm(flight.get("total_duration", sum(s.get("duration", 0) for s in segs))),
        "token": flight.get("departure_token"),
    }

def top_offers(trip, n=3):
    """Top-n cheapest non-stop economy offers (1 search credit)."""
    data = search(trip)
    flights = data.get("best_flights", []) + data.get("other_flights", [])
    flights = [f for f in flights if "price" in f and len(f.get("flights", [])) == 1]
    flights.sort(key=lambda f: f["price"])
    url = data.get("search_metadata", {}).get("google_flights_url")
    return [summarise(f) for f in flights[:n]], url

def return_leg(trip, token):
    """Return-flight details for the cheapest offer (1 extra search credit, only used on alerts)."""
    try:
        data = search(trip, {"departure_token": token})
        flights = data.get("best_flights", []) + data.get("other_flights", [])
        flights = [f for f in flights if "price" in f and len(f.get("flights", [])) == 1]
        return summarise(min(flights, key=lambda f: f["price"])) if flights else None
    except Exception as e:
        print("return-leg lookup failed:", e)
        return None

# ----------------------------------------------------------------- teams card
def txt(text, **kw):
    return {"type": "TextBlock", "text": text, "wrap": True, **kw}

def flight_line(label, o):
    return txt(
        f"**{label}** {o['from']} → {o['to']} · {o['dep']} → {o['arr']} · {o['duration']} · "
        f"**{o['airline']}** {o['numbers']}",
        spacing="Small", size="Small",
    )

def build_card(trip, kind, offers, cur, old=None, ret=None, url=None):
    best = offers[0]
    style, head = {
        "drop": ("good", "🔻 PRICE DROP"),
        "rise": ("attention", "🔺 PRICE INCREASE"),
        "start": ("accent", "✈️ NOW TRACKING"),
        "sample": ("accent", "🧪 SAMPLE ALERT"),
    }[kind]
    legs = legs_of(trip)
    route = f"{legs[0][0]} ✈ {legs[0][1]}" if trip["type"] != "multicity" else " ✈ ".join([legs[0][0]] + [l[1] for l in legs])
    if trip["type"] == "return":
        dates = f"{nice_date(trip['depart'])}  ⇄  {nice_date(trip['return'])}"
    elif trip["type"] == "oneway":
        dates = nice_date(trip["depart"])
    else:
        dates = "  →  ".join(nice_date(l[2]) for l in legs)
    kind_label = {"return": "Round trip", "oneway": "One way", "multicity": "Multi-city"}[trip["type"]]

    header = [
        txt(head, weight="Bolder", size="Small", isSubtle=True),
        txt(route, weight="Bolder", size="ExtraLarge", spacing="None"),
        txt(f"{dates}  ·  {kind_label}  ·  Economy  ·  Non-stop", size="Small", spacing="Small"),
    ]
    price_cols = [{
        "type": "Column", "width": "auto",
        "items": [txt(money(best["price"], cur), size="ExtraLarge", weight="Bolder", spacing="None"),
                  txt("lowest fare" + (" (total)" if trip["type"] != "oneway" else ""), size="Small", isSubtle=True, spacing="None")],
    }]
    if old is not None:
        diff = best["price"] - old
        arrow = "▼" if diff < 0 else "▲"
        price_cols.append({
            "type": "Column", "width": "stretch", "verticalContentAlignment": "Center",
            "items": [txt(f"{arrow} {money(abs(diff), cur)} ({diff / old:+.1%})", weight="Bolder", horizontalAlignment="Right",
                          color="Good" if diff < 0 else "Attention"),
                      txt(f"was {money(old, cur)}", size="Small", isSubtle=True, horizontalAlignment="Right", spacing="None")],
        })

    rows = [{"type": "ColumnSet", "spacing": "Small", "columns": [
        {"type": "Column", "width": "30px", "items": [txt(["🥇", "🥈", "🥉"][i], spacing="None")]},
        {"type": "Column", "width": "stretch", "items": [
            txt(f"**{o['airline']}** {o['numbers']}", spacing="None"),
            txt(f"{o['dep']} → {o['arr']} · {o['duration']}", size="Small", isSubtle=True, spacing="None")]},
        {"type": "Column", "width": "auto", "items": [txt(money(o["price"], cur), weight="Bolder", horizontalAlignment="Right", spacing="None")]},
    ]} for i, o in enumerate(offers)]

    body = [
        {"type": "Container", "style": style, "bleed": True, "items": header},
        {"type": "ColumnSet", "spacing": "Medium", "columns": price_cols},
        {"type": "Container", "separator": True, "spacing": "Medium", "items":
            [txt("TOP 3 NON-STOP FARES" if trip["type"] != "oneway" else "TOP 3 NON-STOP FARES", weight="Bolder", size="Small", isSubtle=True)] + rows},
    ]
    if kind != "sample" and trip["type"] == "return":
        body.append({"type": "Container", "separator": True, "spacing": "Medium", "items": [
            txt("CHEAPEST ITINERARY", weight="Bolder", size="Small", isSubtle=True),
            flight_line("Out", best),
            flight_line("Back", ret) if ret else txt("Return flight: see Google Flights link", size="Small", isSubtle=True),
        ]})
    body.append(txt(f"Checked {datetime.now().strftime('%d %b %H:%M')} · prices for 1 adult", size="Small", isSubtle=True, spacing="Medium"))

    content = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard", "version": "1.4", "body": body,
        "actions": [{"type": "Action.OpenUrl", "title": "Book on Google Flights", "url": url}] if url else [],
    }
    return {"type": "message", "attachments": [{
        "contentType": "application/vnd.microsoft.card.adaptive", "contentUrl": None, "content": content}]}

def post(card):
    r = requests.post(os.environ["TEAMS_WEBHOOK_URL"], json=card, timeout=30)
    r.raise_for_status()

# ----------------------------------------------------------------- commands
def interval_hours(trip):
    """Adaptive polling. Base interval shrinks as departure approaches; the 00:00-04:00 SGT
    window is checked at 2x the rate (hourly at minimum), the rest of the day at half the base rate."""
    days = (datetime.strptime(legs_of(trip)[0][2], "%Y-%m-%d") - datetime.now()).days
    base = 3 if days > 14 else 2 if days > 3 else 1
    return 1 if NIGHT_START <= datetime.now().hour < NIGHT_END else base * 2

def due(trip, state):
    last = state.get("_last", {}).get(trip["id"])
    return last is None or time.time() - last >= interval_hours(trip) * 3600 - 300

def cmd_check(a):
    trips = load_trips()
    state = read(STATE_FILE, {})
    if not trips:
        print("No trips configured.")
        return
    cur = os.getenv("CURRENCY", "INR")
    for t in trips:
        if not a.force and not due(t, state):
            print(f"[{t['id']}] not due yet (every {interval_hours(t)}h)")
            continue
        if not budget_ok(state):
            print(f"Monthly search cap ({MONTHLY_CAP}) reached; skipping.")
            break
        try:
            offers, url = top_offers(t)
        except Exception as e:
            print(f"[{t['id']}] error: {e}")
            continue
        finally:
            state["_usage"]["count"] += 1
        state.setdefault("_last", {})[t["id"]] = time.time()
        if not offers:
            print(f"[{t['id']}] no non-stop offers found")
            continue
        best = offers[0]
        price = best["price"]
        old = state.get(t["id"])
        print(f"[{t['id']}] {describe(t)}: {price} {cur} via {best['airline']} (prev {old})")
        for o in offers:
            print(f"   {money(o['price'], cur)}  {o['airline']} {o['numbers']}  {o['dep']} -> {o['arr']}")
        if old is None or price != old:
            ret = None
            if t["type"] == "return" and best.get("token") and budget_ok(state):
                ret = return_leg(t, best["token"])
                state["_usage"]["count"] += 1
            kind = "start" if old is None else ("drop" if price < old else "rise")
            post(build_card(t, kind, offers, cur, old=old, ret=ret, url=url))
        state[t["id"]] = price
        time.sleep(1)
    write(STATE_FILE, state)

def cmd_test(_):
    trip = {"type": "return", "from": "SIN", "to": "BOM", "depart": "2026-11-24", "return": "2027-01-09"}
    mk = lambda p, a, n, d, e: {"price": p, "airline": a, "numbers": n, "dep": d, "arr": e, "from": "SIN", "to": "BOM", "duration": "5h 35m"}
    offers = [mk(31200, "IndiGo", "6E 1006", "Tue 24 Nov, 23:55", "Wed 25 Nov, 02:50"),
              mk(33850, "Singapore Airlines", "SQ 424", "Tue 24 Nov, 08:15", "Tue 24 Nov, 11:20"),
              mk(34900, "Air India", "AI 2381", "Tue 24 Nov, 14:00", "Tue 24 Nov, 17:05")]
    post(build_card(trip, "sample", offers, "INR", old=33400, url="https://www.google.com/travel/flights"))

def cmd_add(a):
    trips = read(TRIPS_FILE, [])
    tid = str(max([int(t["id"]) for t in trips if str(t.get("id", "")).isdigit()], default=0) + 1)
    if a.type == "multicity":
        if not a.legs or len(a.legs) < 2:
            sys.exit("multicity needs --legs ORG-DST:YYYY-MM-DD (2+)")
        legs = []
        for l in a.legs:
            route, date = l.split(":")
            o, d = route.split("-")
            legs.append([o.upper(), d.upper(), date])
        trip = {"id": tid, "type": "multicity", "legs": legs}
    else:
        if not (a.origin and a.to and a.depart):
            sys.exit("--from, --to and --depart are required")
        trip = {"id": tid, "type": a.type, "from": a.origin.upper(), "to": a.to.upper(), "depart": a.depart}
        if a.type == "return":
            if not a.ret:
                sys.exit("--return date required for return trips")
            trip["return"] = a.ret
    trips.append(trip)
    write(TRIPS_FILE, trips)
    print(f"Added trip {tid}: {describe(trip)}")

def cmd_list(_):
    for t in load_trips():
        print(t["id"], describe(t))

def cmd_remove(a):
    write(TRIPS_FILE, [t for t in load_trips() if t["id"] != a.id])

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # ₹ / emoji safe on Windows consoles
    load_env()
    p = argparse.ArgumentParser()
    s = p.add_subparsers(dest="cmd", required=True)
    add = s.add_parser("add")
    add.add_argument("--type", choices=["return", "oneway", "multicity"], required=True)
    add.add_argument("--from", dest="origin")
    add.add_argument("--to")
    add.add_argument("--depart")
    add.add_argument("--return", dest="ret")
    add.add_argument("--legs", nargs="+")
    add.set_defaults(f=cmd_add)
    s.add_parser("list").set_defaults(f=cmd_list)
    rm = s.add_parser("remove"); rm.add_argument("id"); rm.set_defaults(f=cmd_remove)
    ck = s.add_parser("check"); ck.add_argument("--force", action="store_true"); ck.set_defaults(f=cmd_check)
    s.add_parser("test").set_defaults(f=cmd_test)
    a = p.parse_args()
    a.f(a)

if __name__ == "__main__":
    main()
