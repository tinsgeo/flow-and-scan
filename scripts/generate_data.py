#!/usr/bin/env python3
"""
Generates the two plain-HTML data tables used for VBA/Excel practice:
  - data/associates.html
  - data/orders.html

These pages are intentionally plain HTML tables with NO JavaScript.
That's what makes them readable by Excel's "Data > From Web" / VBA
QueryTables web queries, which fetch raw HTML and do not run scripts.

Run manually with:  python scripts/generate_data.py
Run automatically by .github/workflows/update-data.yml on a schedule.
"""

import random
import datetime
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")

POSITIONS = ["Pick", "Pack", "Sort", "Induct", "Stow"]
FIRST_NAMES = ["Jordan", "Casey", "Morgan", "Taylor", "Riley", "Avery", "Devon",
               "Reese", "Sydney", "Cameron", "Peyton", "Quinn", "Rowan",
               "Emerson", "Harper", "Elliot"]
LAST_NAMES = ["Bell", "Reyes", "Whitfield", "Nakamura", "Alvarez", "Sutton",
              "Okafor", "Petrov", "Lindgren", "Marsh", "Delgado", "Osei",
              "Kowalski", "Fontaine"]

PAGE_HEAD = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{title} — Data Table (Flow &amp; Scan)</title>
<style>
  body {{ font-family: 'IBM Plex Mono', ui-monospace, monospace; background:#10141c; color:#e9edf5; margin:0; padding:24px 20px 60px; }}
  h1 {{ font-size:20px; margin-bottom:4px; }}
  p {{ color:#8891a6; font-size:13px; max-width:70ch; }}
  a {{ color:#4fd6e8; }}
  table {{ border-collapse: collapse; width:100%; margin-top:18px; font-size:13px; }}
  th, td {{ text-align:left; padding:6px 10px; border-bottom:1px solid #2a3242; }}
  th {{ color:#8891a6; text-transform:uppercase; font-size:11px; letter-spacing:0.04em; }}
  td.num, th.num {{ text-align:right; }}
</style>
</head>
<body>
<h1>{title}</h1>
<p>Generated {timestamp} UTC. Synthetic practice data — regenerates on a fixed
schedule. Plain HTML table, no JavaScript, safe to pull with Excel's
"Data &gt; From Web" or a VBA QueryTables web query. Back to
<a href="../index.html">Flow &amp; Scan home</a>.</p>
"""

PAGE_TAIL = """
</body>
</html>
"""


def fmt_time(total_minutes):
    h = int(total_minutes // 60)
    m = int(total_minutes % 60)
    ampm = "PM" if h >= 12 else "AM"
    h12 = h % 12
    if h12 == 0:
        h12 = 12
    return f"{h12}:{m:02d} {ampm}"


def generate_associates_table(rows_target=90):
    shift_start = 6 * 60
    shift_end = 14.5 * 60
    rows = []
    associate_count = 16
    for i in range(associate_count):
        assoc_id = 1000 + i
        name = f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"
        position = random.choice(POSITIONS)
        t = shift_start + random.randint(0, 10)
        while t < shift_end:
            interval = random.randint(2, 8)
            if random.random() < 0.12:
                interval = random.randint(18, 44)
            t += interval
            if t >= shift_end:
                break
            unit = f"SKU-{random.randint(10000, 99999)}"
            rows.append((assoc_id, name, position, fmt_time(t), unit))

    rows.sort(key=lambda r: (r[0], r[3]))

    html = [PAGE_HEAD.format(
        title="Associate Scan Data",
        timestamp=datetime.datetime.now(datetime.UTC).strftime("%Y-%m-%d %H:%M"),
    )]
    html.append("<table>")
    html.append("<thead><tr><th>ID</th><th>Name</th><th>Position</th>"
                 "<th>Scan Time</th><th>Unit Scanned</th></tr></thead>")
    html.append("<tbody>")
    for assoc_id, name, position, scan_time, unit in rows:
        html.append(
            f"<tr><td>{assoc_id}</td><td>{name}</td><td>{position}</td>"
            f"<td>{scan_time}</td><td>{unit}</td></tr>"
        )
    html.append("</tbody></table>")
    html.append(PAGE_TAIL)
    return "\n".join(html)


def generate_orders_table(entries=40):
    pending, created, in_progress, completed = 40, 0, 6, 0
    now = datetime.datetime.now(datetime.UTC)
    rows = []
    for i in range(entries):
        ts = now - datetime.timedelta(minutes=(entries - i) * 2)
        new_orders = random.randint(0, 4)
        created += new_orders
        pending += new_orders

        started = min(pending, random.randint(0, 3))
        pending -= started
        in_progress += started

        finished = min(in_progress, random.randint(0, 3))
        in_progress -= finished
        completed += finished

        rows.append((ts.strftime("%H:%M:%S"), pending, created, in_progress, completed))

    html = [PAGE_HEAD.format(
        title="Order Flow Data",
        timestamp=now.strftime("%Y-%m-%d %H:%M"),
    )]
    html.append("<table>")
    html.append("<thead><tr><th>Timestamp (UTC)</th><th class='num'>Pending</th>"
                 "<th class='num'>Created</th><th class='num'>In Progress</th>"
                 "<th class='num'>Completed</th></tr></thead>")
    html.append("<tbody>")
    for ts, p, c, ip, comp in rows:
        html.append(
            f"<tr><td>{ts}</td><td class='num'>{p}</td><td class='num'>{c}</td>"
            f"<td class='num'>{ip}</td><td class='num'>{comp}</td></tr>"
        )
    html.append("</tbody></table>")
    html.append(PAGE_TAIL)
    return "\n".join(html)


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(os.path.join(DATA_DIR, "associates.html"), "w", encoding="utf-8") as f:
        f.write(generate_associates_table())
    with open(os.path.join(DATA_DIR, "orders.html"), "w", encoding="utf-8") as f:
        f.write(generate_orders_table())
    print("Generated data/associates.html and data/orders.html")


if __name__ == "__main__":
    main()
