# Drafter Draftsheet

Local app that exports one or more Drafter.lol drafts to a formatted Excel workbook.

## Requirements

- Python 3.10 or newer
- Node.js 20.19+ or 22.12+
- Microsoft Edge for local exports. The Render container uses Chromium under a virtual display.

Install the Python and frontend dependencies from the project root:

```powershell
python -m pip install -r requirements.txt
npm --prefix web install
```

## Run

Start the API in one terminal:

```powershell
python -m uvicorn api:app --host 127.0.0.1 --port 8001
```

Start the React app in another terminal:

```powershell
npm --prefix web run dev -- --host 127.0.0.1
```

Open the URL printed by Vite, usually `http://127.0.0.1:5173`. The API also exposes interactive docs at `http://127.0.0.1:8001/docs`.

The API opens a visible Edge window while reading drafts because Drafter.lol rejects headless requests. Paste up to 50 Drafter.lol draft URLs, one per line. All drafts are written as labeled blocks in a single Excel worksheet, with champion names and thumbnails for every pick and ban.

## Deploy on Render

The root `render.yaml` defines a public Docker web service. To deploy it, push this project to a GitHub repository, then create a Blueprint in the Render Dashboard and connect that repository. Render builds the frontend and API together and assigns an `onrender.com` URL.

The Blueprint uses Render's Free plan to avoid opting into paid compute. Free services sleep after 15 minutes without traffic and can take about a minute to wake. The plan provides 512 MB of RAM; Chromium may require a larger paid plan for reliable exports. Check the service's resource usage before upgrading.