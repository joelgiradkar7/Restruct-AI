# RESTRUCT AI

AI-assisted structural damage & rehabilitation advisor. Users describe a damaged
structural component, optionally attach a site photograph, and receive an
evidence-based preliminary assessment plus three ranked rehabilitation pathways
(materials, procedure, safety, IS-code considerations). Reports can be exported
as PDF, DOCX, or XLSX.

The AI is instructed to clearly separate **observed / provided** facts from
**estimated / inferred** possibilities and from what **cannot be determined**
from the available evidence. It will not fabricate crack widths, capacities, or
hidden reinforcement condition.

Supported damage categories include cracking, spalling, delamination,
reinforcement corrosion, fire / thermal, earthquake / seismic, blast /
explosion, vehicle impact, water / moisture, chemical deterioration,
differential settlement, foundation distress, overloading, construction
defects, material deterioration, structural deformation, masonry damage,
connection / joint damage, and unknown / other.

## Technology stack

| Layer     | Technology                                                                |
| --------- | ------------------------------------------------------------------------- |
| Frontend  | React 19 (Create React App via CRACO), Axios, lucide-react, Tailwind CSS  |
| Backend   | FastAPI, Uvicorn, Motor (async MongoDB), Pydantic v2                      |
| Database  | MongoDB                                                                   |
| AI        | OpenAI GPT-5.4 Mini via `emergentintegrations` (Emergent Universal LLM key) |
| Reports   | ReportLab (PDF), python-docx (DOCX), openpyxl (XLSX)                      |
| Imaging   | Pillow (server-side image validation)                                     |
| Auth      | None (no user accounts in the current build)                              |

## Project structure

```
/app
├── README.md
├── design_guidelines.json
├── backend/
│   ├── server.py               # FastAPI app: routes, AI prompt, normalization, exports
│   ├── requirements.txt
│   ├── pytest.ini
│   ├── .env                    # NOT committed (real values)
│   ├── .env.example            # Variable names only
│   └── tests/
│       └── test_assessment_live.py
├── frontend/
│   ├── package.json
│   ├── craco.config.js
│   ├── tailwind.config.js
│   ├── postcss.config.js
│   ├── components.json
│   ├── public/
│   ├── src/
│   │   ├── index.js
│   │   ├── App.js              # AssessmentForm, EvidencePanel, Results, History
│   │   ├── App.css
│   │   ├── index.css
│   │   ├── components/ui/      # shadcn/ui primitives
│   │   └── hooks/
│   ├── .env                    # NOT committed (real values)
│   └── .env.example            # Variable names only
├── memory/
│   └── PRD.md
├── tests/
└── test_reports/
```

## Required runtime versions

| Component     | Version                                                   |
| ------------- | --------------------------------------------------------- |
| Node.js       | 18 LTS or 20 LTS (React 19 + react-scripts 5)             |
| Yarn          | 1.22.x (Classic) — declared in `packageManager`           |
| Python        | 3.11 (tested); 3.10+ is expected to work                  |
| MongoDB       | 5.0 or newer (local or MongoDB Atlas)                     |
| OS            | Windows 11 (this guide), macOS, or Linux                  |

## Environment variables

### `backend/.env`

| Variable            | Purpose                                                     |
| ------------------- | ----------------------------------------------------------- |
| `MONGO_URL`         | MongoDB connection string (e.g. `mongodb://localhost:27017`) |
| `DB_NAME`           | Database name to use                                        |
| `CORS_ORIGINS`      | Comma-separated allowed origins (use `*` for local dev)     |
| `EMERGENT_LLM_KEY`  | Emergent Universal LLM key (required for AI assessments)    |

### `frontend/.env`

| Variable                | Purpose                                                                 |
| ----------------------- | ----------------------------------------------------------------------- |
| `REACT_APP_BACKEND_URL` | Base URL of the backend (e.g. `http://localhost:8001`). All frontend API calls append `/api`. |
| `WDS_SOCKET_PORT`       | Webpack dev-server socket port. Set to `443` behind an HTTPS proxy or leave `0` locally. |
| `ENABLE_HEALTH_CHECK`   | `true` / `false` — dev-server health-check toggle                       |

Sample `.env.example` files are provided in both `backend/` and `frontend/`.
Never commit real `.env` files (they are ignored by `.gitignore` patterns).

## AI / external services

- **Emergent Universal LLM key** — one key that works across the OpenAI /
  Anthropic / Gemini providers exposed by `emergentintegrations`. Currently the
  backend requests `openai / gpt-5.4-mini`.
- **MongoDB** — used only for persisting assessments and simple status checks.

No third-party authentication, payments, or storage services are used in this
build.

## Installation (local Windows 11)

### Step 1 — Install required software

- **Git for Windows** — <https://git-scm.com/download/win>
- **Node.js 20 LTS** (includes npm) — <https://nodejs.org/en/download>
- **Yarn Classic 1.22.x** (after Node is installed):
  ```powershell
  corepack enable
  corepack prepare yarn@1.22.22 --activate
  ```
- **Python 3.11** — <https://www.python.org/downloads/windows/> (tick “Add python.exe to PATH”)
- **MongoDB Community Server 7.x** — <https://www.mongodb.com/try/download/community> (default port `27017`) or a free MongoDB Atlas cluster.

### Step 2 — Obtain the project

```powershell
git clone <your-repo-url> restruct-ai
cd restruct-ai
```

### Step 3 — Install dependencies

Backend:

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
cd ..
```

Frontend:

```powershell
cd frontend
yarn install
cd ..
```

### Step 4 — Configure `.env`

```powershell
copy backend\.env.example backend\.env
copy frontend\.env.example frontend\.env
```

Fill in real values (minimum for local use):

```
# backend/.env
MONGO_URL=mongodb://localhost:27017
DB_NAME=restruct_ai
CORS_ORIGINS=*
EMERGENT_LLM_KEY=<paste your Emergent Universal LLM key>

# frontend/.env
REACT_APP_BACKEND_URL=http://localhost:8001
WDS_SOCKET_PORT=0
ENABLE_HEALTH_CHECK=false
```

### Step 5 — Configure the database

No migrations. Motor / PyMongo creates the `restruct_ai` database and the
`assessments` / `status_checks` collections on first write. Start MongoDB:

```powershell
net start MongoDB
```

### Step 6 — Configure the AI service

Paste the Emergent Universal LLM key into `backend/.env` as `EMERGENT_LLM_KEY`.
If missing, `POST /api/assessments` returns HTTP 503.

### Step 7 — Migrations / setup

None required.

### Step 8 — Start the backend

```powershell
cd backend
.\.venv\Scripts\activate
uvicorn server:app --host 0.0.0.0 --port 8001 --reload
```

Health check: <http://localhost:8001/api/> → `{"message":"ResructAI API ready"}`.

### Step 9 — Start the frontend

In a separate terminal:

```powershell
cd frontend
yarn start
```

CRACO / react-scripts serve the app at <http://localhost:3000>.

### Step 10 — Open the app

Visit <http://localhost:3000>. The browser calls
`http://localhost:8001/api/assessments`.

## Production build (frontend)

```powershell
cd frontend
yarn build
```

Output is written to `frontend/build/`. Point `REACT_APP_BACKEND_URL` at your
production backend before building. The backend has no build step — deploy
`server.py` + `requirements.txt` and run under a production ASGI server
(e.g. `uvicorn server:app --host 0.0.0.0 --port 8001 --workers 2`).

## Default local ports

| Service   | Port  |
| --------- | ----- |
| Frontend  | 3000  |
| Backend   | 8001  |
| MongoDB   | 27017 |

## Troubleshooting

- **`{"detail":"AI service is not configured"}`** — `EMERGENT_LLM_KEY` is
  missing or empty in `backend/.env`. Restart the backend after adding it.
- **`Only JPG, PNG, or WEBP photos are supported`** — server-side image
  validation rejects other MIME types.
- **`Please upload a photo with visible structural detail`** — the uploaded
  image is uniform / blank; Pillow’s min/max luminance check rejects it.
- **CORS errors in the browser** — set `CORS_ORIGINS=*` (or your exact
  frontend origin) in `backend/.env` and restart the backend.
- **Frontend cannot reach the backend** — confirm `REACT_APP_BACKEND_URL` is
  a full URL like `http://localhost:8001`, then restart `yarn start`.
- **`npm` warnings from tooling** — use **`yarn`**, not `npm`, for the
  frontend; `package.json` and lockfile are Yarn 1 Classic.
- **MongoDB refuses connections** — start the service (`net start MongoDB`)
  or update `MONGO_URL` to your Atlas SRV string.

## API summary

| Method | Path                                              | Purpose                                |
| ------ | ------------------------------------------------- | -------------------------------------- |
| GET    | `/api/`                                           | Health check                           |
| POST   | `/api/assessments`                                | Create a live AI assessment            |
| GET    | `/api/assessments`                                | List saved assessments                 |
| GET    | `/api/assessments/{id}/report/{pdf\|docx\|xlsx}` | Download report in the chosen format   |
| POST   | `/api/status`                                     | Insert a status check                  |
| GET    | `/api/status`                                     | List status checks                     |

All frontend calls go through `REACT_APP_BACKEND_URL + "/api"`.

## Files to back up

Essential project files (exclude `node_modules/`, `.venv/`, build output, and
any real `.env`):

- `README.md`
- `.gitignore`
- `design_guidelines.json`
- `backend/server.py`
- `backend/requirements.txt`
- `backend/pytest.ini`
- `backend/.env.example`
- `backend/tests/`
- `frontend/package.json`
- `frontend/yarn.lock`
- `frontend/craco.config.js`
- `frontend/tailwind.config.js`
- `frontend/postcss.config.js`
- `frontend/components.json`
- `frontend/jsconfig.json`
- `frontend/public/`
- `frontend/src/`
- `frontend/.env.example`
- `memory/PRD.md`

MongoDB data is not stored in the repo. Back it up separately with
`mongodump`.

## GitHub

There is no GitHub integration configured in this project. To back the code up
to a new repository:

```powershell
cd restruct-ai
git init
git add .
git commit -m "RESTRUCT AI initial backup"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

`.gitignore` already excludes `node_modules/`, virtualenvs, build output, and
credential-shaped files. Confirm that neither `backend/.env` nor
`frontend/.env` is committed before pushing.
