# Flow & Scan — Practice Terminal

A companion practice site for *The VBA Portfolio Project*. Readers use this
site to practice two skills from the book:

1. **Seeing** operational data visually (the productivity and order-flow
   dashboards)
2. **Pulling** that same kind of data into Excel with VBA (the plain data
   tables under `/data/`)

## What's in this repo

```
index.html            landing page
productivity.html     visual dashboard — associate scans, gaps, filters
orders.html           visual dashboard — live order-flow counters
styles.css            shared styling
app.js                shared JS for the visual dashboards
data/
  associates.html     plain HTML table — VBA/Excel-scrapable, no JS
  orders.html          plain HTML table — VBA/Excel-scrapable, no JS
scripts/
  generate_data.py    regenerates the two files in /data/
.github/workflows/
  update-data.yml     runs generate_data.py every 30 minutes and commits
```

## Why two kinds of pages

`productivity.html` and `orders.html` are built with JavaScript so they're
interactive to look at in a browser. But VBA's built-in web-query tools
(`QueryTables.Add`, or Excel's Data > From Web) fetch raw HTML — they don't
run JavaScript. So the pages under `/data/` exist specifically as **plain
HTML tables** with no script involved, which is what a web query can actually
read. Those two files are regenerated automatically on a schedule by the
GitHub Action in this repo, so the data keeps changing over time even though
GitHub Pages itself only serves static files.

## Deploying to GitHub Pages

1. Create a new repository and push this folder's contents to it.
2. In the repo, go to **Settings → Pages**.
3. Under "Build and deployment," set **Source** to `Deploy from a branch`,
   branch `main`, folder `/ (root)`. Save.
4. GitHub will give you a URL like `https://tinsgeo.github.io/flow-and-scan/`.
   That's the link to share in the book.
5. Go to the **Actions** tab and confirm the "Regenerate practice data"
   workflow is enabled. You can also trigger it manually the first time
   with the "Run workflow" button so the data isn't stale from the initial
   commit.

No build step, no server, no dependencies beyond what's already in this repo.

## Note for the book

When explaining the VBA side to readers, the simplest approach is Excel's
built-in web query, pointed at the `/data/` pages directly, e.g.:

```vba
Sub PullAssociateData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("AssociateData")

    With ws.QueryTables.Add(Connection:="URL;https://tinsgeo.github.io/flow-and-scan/data/associates.html", _
        Destination:=ws.Range("A1"))
        .RefreshOnFileOpen = False
        .Refresh BackgroundQuery:=False
    End With
End Sub
```

That pulls the whole table straight into a sheet — from there, readers apply
what they learned in Chapter 2 (loops, `DateDiff`, `If`) to calculate gaps or
build their own flow view, exactly like the case studies in Chapters 3 and 4.
