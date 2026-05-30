import os
import uuid
import shutil
import json
import logging
from datetime import datetime, timedelta
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

# Load environmental configurations
load_dotenv()

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Lazy import to avoid circular dependencies
from trackon_automation import TrackonAutomation

app = FastAPI(
    title="Trackon API & Automation Portal",
    description="Exposes Trackon courier weights automation as a high-performance REST API with headless CAPTCHA solving.",
    version="1.0.0"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global in-memory registry of active runs (in addition to file persistence)
active_runs = {}

RUNS_DIR = os.path.join(os.getcwd(), "runs")
os.makedirs(RUNS_DIR, exist_ok=True)

def run_automation_task(run_id: str, excel_path: str):
    """Background task executor for sheet updates."""
    try:
        username = os.getenv("TRACKON_USERNAME")
        password = os.getenv("TRACKON_PASSWORD")
        captcha_key = os.getenv("TWOCAPTCHA_API_KEY")
        
        # Determine headless execution based on 2Captcha key availability
        headless = bool(username and password and captcha_key)
        
        runner = TrackonAutomation(
            excel_path=excel_path,
            username=username,
            password=password,
            captcha_api_key=captcha_key,
            headless=headless,
            run_id=run_id
        )
        
        active_runs[run_id] = runner
        runner.run()
        
    except Exception as e:
        logger.error(f"Background run {run_id} failed: {e}")
        # Update metadata to failed state manually if runner crashed before creating runner
        run_dir = os.path.join(RUNS_DIR, run_id)
        os.makedirs(run_dir, exist_ok=True)
        failed_state = {
            "run_id": run_id,
            "status": "failed",
            "progress": 0,
            "error": str(e),
            "logs": [f"Fatal system error: {e}"],
            "end_time": datetime.now().isoformat()
        }
        with open(os.path.join(run_dir, "metadata.json"), "w") as f:
            json.dump(failed_state, f, indent=4)

@app.post("/api/upload")
async def upload_file(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Upload bookings.xlsx and kick off the automation run in the background."""
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only standard .xlsx spreadsheets are supported.")
        
    run_id = uuid.uuid4().hex[:12]
    run_dir = os.path.join(RUNS_DIR, run_id)
    os.makedirs(run_dir, exist_ok=True)
    
    excel_path = os.path.join(run_dir, "bookings.xlsx")
    with open(excel_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    # Queue the automation run as a background task
    background_tasks.add_task(run_automation_task, run_id, excel_path)
    
    return JSONResponse(
        content={
            "run_id": run_id,
            "status": "queued",
            "message": "Automation successfully triggered in the background. Use the status endpoint to monitor progress.",
            "status_url": f"/api/status/{run_id}"
        },
        status_code=202
    )

@app.get("/api/status/{run_id}")
async def get_run_status(run_id: str):
    """Check the progress, logs, and status of an active or historical run."""
    run_dir = os.path.join(RUNS_DIR, run_id)
    metadata_path = os.path.join(run_dir, "metadata.json")
    
    # 1. First, check active in-memory runners (real-time updates)
    if run_id in active_runs:
        return active_runs[run_id].state
        
    # 2. Fall back to persisted metadata on disk
    if os.path.exists(metadata_path):
        try:
            with open(metadata_path, "r") as f:
                data = json.load(f)
            # Add download URL if completed
            if data.get("status") == "completed":
                data["download_url"] = f"/api/download/{run_id}"
            return data
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Failed to read run metadata: {e}")
            
    raise HTTPException(status_code=404, detail="Run ID not found.")

@app.get("/api/download/{run_id}")
async def download_run_result(run_id: str):
    """Downloads the completed populated spreadsheet."""
    run_dir = os.path.join(RUNS_DIR, run_id)
    result_path = os.path.join(run_dir, "bookings_updated.xlsx")
    
    if os.path.exists(result_path):
        return FileResponse(
            path=result_path,
            filename=f"bookings_updated_{run_id}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    # Check if run is still running or failed
    metadata_path = os.path.join(run_dir, "metadata.json")
    if os.path.exists(metadata_path):
        with open(metadata_path, "r") as f:
            data = json.load(f)
        status = data.get("status")
        if status == "running":
            raise HTTPException(status_code=400, detail="Excel generation is still in progress. Please wait.")
        elif status == "failed":
            raise HTTPException(status_code=400, detail=f"Excel generation failed. Error: {data.get('error')}")
            
    raise HTTPException(status_code=404, detail="Result file not found.")

@app.get("/api/query/{awb}")
async def query_single_awb(
    awb: str,
    from_date: Optional[str] = Query(None, description="Start date (YYYY-MM-DD)"),
    to_date: Optional[str] = Query(None, description="End date (YYYY-MM-DD)")
):
    """Queries a single AWB on-demand and returns its weight directly."""
    username = os.getenv("TRACKON_USERNAME")
    password = os.getenv("TRACKON_PASSWORD")
    captcha_key = os.getenv("TWOCAPTCHA_API_KEY")
    
    if not username or not password:
        raise HTTPException(status_code=400, detail="Franchisee credentials not configured in system.")
        
    # Default range to past 7 days if not provided
    if not to_date:
        to_dt = datetime.now()
    else:
        to_dt = datetime.strptime(to_date, "%Y-%m-%d")
        
    if not from_date:
        from_dt = to_dt - timedelta(days=6)
    else:
        from_dt = datetime.strptime(from_date, "%Y-%m-%d")
        
    # Max range of 7 days check for single search batch
    if (to_dt - from_dt).days > 6:
        raise HTTPException(status_code=400, detail="Trackon only supports a maximum range of 7 days per search query.")

    logger.info(f"Triggering on-demand query for AWB: {awb} in range {from_dt.strftime('%d-%m-%Y')} to {to_dt.strftime('%d-%m-%Y')}")
    
    headless = bool(captcha_key)
    runner = TrackonAutomation(
        excel_path="bookings.xlsx", # Dummy, not used
        username=username,
        password=password,
        captcha_api_key=captcha_key,
        headless=headless
    )
    
    try:
        runner.setup_driver()
        runner.login_manual()
        runner.navigate_to_reports()
        runner.perform_search(from_dt, to_dt)
        results = runner.extract_table_data()
        
        awb_cleaned = str(awb).strip()
        weight = results.get(awb_cleaned)
        
        if weight:
            # Clean and format
            cleaned_w = str(weight).upper().replace("KG", "").strip()
            formatted_w = f"{cleaned_w} KG"
            return {"awb": awb, "weight": formatted_w, "found": True}
        else:
            # Try substring matching
            matched_key = next((k for k in results if awb_cleaned in k or k in awb_cleaned), None)
            if matched_key:
                cleaned_w = str(results[matched_key]).upper().replace("KG", "").strip()
                formatted_w = f"{cleaned_w} KG"
                return {"awb": awb, "weight": formatted_w, "found": True, "matched_docket": matched_key}
                
            return {"awb": awb, "weight": None, "found": False, "message": "AWB not found in report records for this range."}
            
    except Exception as e:
        logger.error(f"Single AWB query failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if runner.driver:
            runner.driver.quit()

@app.get("/", response_class=HTMLResponse)
async def landing_page():
    """Beautiful, dark-mode, glassmorphic UI landing page for active run stats and uploads."""
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Trackon Courier Weight Automation Portal</title>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            :root {
                --bg: #090a0f;
                --card-bg: rgba(255, 255, 255, 0.03);
                --card-border: rgba(255, 255, 255, 0.08);
                --primary: #6366f1;
                --primary-hover: #4f46e5;
                --text: #e2e8f0;
                --text-muted: #94a3b8;
                --success: #10b981;
            }
            * {
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }
            body {
                background-color: var(--bg);
                color: var(--text);
                font-family: 'Plus Jakarta Sans', sans-serif;
                min-height: 100vh;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                overflow-x: hidden;
                position: relative;
            }
            /* Vibrant glowing background blobs */
            .blob {
                position: absolute;
                width: 500px;
                height: 500px;
                background: radial-gradient(circle, rgba(99, 102, 241, 0.15) 0%, rgba(99, 102, 241, 0) 70%);
                border-radius: 50%;
                z-index: -1;
                filter: blur(40px);
            }
            .blob-1 { top: -10%; left: -10%; }
            .blob-2 { bottom: -10%; right: -10%; }

            .container {
                max-width: 800px;
                width: 90%;
                padding: 40px;
                background: var(--card-bg);
                border: 1px solid var(--card-border);
                border-radius: 24px;
                backdrop-filter: blur(20px);
                box-shadow: 0 20px 40px rgba(0, 0, 0, 0.5);
                text-align: center;
            }
            h1 {
                font-size: 2.5rem;
                font-weight: 700;
                background: linear-gradient(135deg, #fff 0%, #a5b4fc 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin-bottom: 12px;
            }
            .subtitle {
                font-size: 1.1rem;
                color: var(--text-muted);
                margin-bottom: 40px;
            }
            .status-badge {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                background: rgba(16, 185, 129, 0.1);
                border: 1px solid rgba(16, 185, 129, 0.2);
                color: var(--success);
                padding: 6px 16px;
                border-radius: 100px;
                font-weight: 600;
                font-size: 0.85rem;
                margin-bottom: 24px;
            }
            .status-badge::before {
                content: '';
                width: 8px;
                height: 8px;
                background: var(--success);
                border-radius: 50%;
                box-shadow: 0 0 10px var(--success);
            }
            /* Drag and Drop upload area */
            .upload-area {
                border: 2px dashed var(--card-border);
                background: rgba(255, 255, 255, 0.01);
                padding: 40px 20px;
                border-radius: 16px;
                cursor: pointer;
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                margin-bottom: 30px;
                position: relative;
            }
            .upload-area:hover {
                border-color: var(--primary);
                background: rgba(99, 102, 241, 0.02);
            }
            .upload-icon {
                font-size: 3rem;
                margin-bottom: 16px;
                display: block;
            }
            .upload-text {
                font-weight: 600;
                margin-bottom: 6px;
            }
            .upload-subtext {
                font-size: 0.85rem;
                color: var(--text-muted);
            }
            input[type="file"] {
                display: none;
            }
            /* Button styling */
            .btn {
                background: var(--primary);
                color: white;
                border: none;
                padding: 14px 28px;
                border-radius: 12px;
                font-weight: 600;
                cursor: pointer;
                font-size: 1rem;
                transition: background 0.2s ease, transform 0.1s ease;
                display: inline-flex;
                align-items: center;
                gap: 8px;
                box-shadow: 0 4px 14px rgba(99, 102, 241, 0.4);
            }
            .btn:hover {
                background: var(--primary-hover);
                transform: translateY(-1px);
            }
            .btn:active {
                transform: translateY(0);
            }
            
            /* Status box */
            .status-box {
                margin-top: 30px;
                background: rgba(255, 255, 255, 0.01);
                border: 1px solid var(--card-border);
                border-radius: 16px;
                padding: 24px;
                display: none;
                text-align: left;
            }
            .progress-bar-container {
                width: 100%;
                background: rgba(255, 255, 255, 0.05);
                border-radius: 100px;
                height: 8px;
                margin: 16px 0;
                overflow: hidden;
            }
            .progress-bar {
                width: 0%;
                height: 100%;
                background: var(--primary);
                box-shadow: 0 0 10px var(--primary);
                border-radius: 100px;
                transition: width 0.4s ease;
            }
            .log-box {
                background: #020306;
                border: 1px solid var(--card-border);
                font-family: monospace;
                font-size: 0.8rem;
                padding: 12px;
                border-radius: 8px;
                height: 150px;
                overflow-y: auto;
                color: #a5b4fc;
                margin-top: 14px;
            }
        </style>
    </head>
    <body>
        <div class="blob blob-1"></div>
        <div class="blob blob-2"></div>
        
        <div class="container">
            <span class="status-badge">API Running Headlessly</span>
            <h1>Trackon Weight Sync Engine</h1>
            <p class="subtitle">Extract shipment weights from your franchisee portal and update Excel sheets instantly.</p>
            
            <form id="uploadForm">
                <label for="excelFile" class="upload-area" id="dropZone">
                    <span class="upload-icon">📂</span>
                    <span class="upload-text" id="uploadText">Drag & Drop bookings.xlsx here</span>
                    <span class="upload-subtext">or click to browse your computer</span>
                    <input type="file" id="excelFile" name="file" accept=".xlsx">
                </label>
                <button type="submit" class="btn" id="submitBtn">🚀 Start Sync Task</button>
            </form>
            
            <div class="status-box" id="statusBox">
                <h3 id="statusTitle" style="font-weight:600; display:flex; justify-content:space-between;">
                    <span>Sync Progress</span>
                    <span id="percentText" style="color:var(--primary)">0%</span>
                </h3>
                <div class="progress-bar-container">
                    <div class="progress-bar" id="progressBar"></div>
                </div>
                <div style="font-size:0.9rem; color:var(--text-muted); display:flex; justify-content:space-between; margin-bottom:8px;">
                    <span id="rowsCount">Processed: 0/0 rows</span>
                    <span id="syncStatus">Status: Queued</span>
                </div>
                <div class="log-box" id="logBox"></div>
                <div id="downloadContainer" style="margin-top:16px; text-align:center; display:none;">
                    <a id="downloadBtn" href="#" class="btn" style="background:var(--success); box-shadow:0 4px 14px rgba(16,185,129,0.3)">📥 Download Updated Excel</a>
                </div>
            </div>
        </div>

        <script>
            const form = document.getElementById("uploadForm");
            const fileInput = document.getElementById("excelFile");
            const dropZone = document.getElementById("dropZone");
            const submitBtn = document.getElementById("submitBtn");
            
            const statusBox = document.getElementById("statusBox");
            const progressBar = document.getElementById("progressBar");
            const percentText = document.getElementById("percentText");
            const rowsCount = document.getElementById("rowsCount");
            const syncStatus = document.getElementById("syncStatus");
            const logBox = document.getElementById("logBox");
            const downloadContainer = document.getElementById("downloadContainer");
            const downloadBtn = document.getElementById("downloadBtn");
            const uploadText = document.getElementById("uploadText");
            
            let pollInterval = null;

            // Handle filename text update
            fileInput.addEventListener("change", () => {
                if (fileInput.files.length > 0) {
                    uploadText.textContent = fileInput.files[0].name;
                }
            });

            form.addEventListener("submit", async (e) => {
                e.preventDefault();
                if (fileInput.files.length === 0) {
                    alert("Please upload an Excel spreadsheet (.xlsx) file first.");
                    return;
                }
                
                const formData = new FormData();
                formData.append("file", fileInput.files[0]);
                
                submitBtn.disabled = true;
                submitBtn.textContent = "Uploading...";
                
                try {
                    const response = await fetch("/api/upload", {
                        method: "POST",
                        body: formData
                    });
                    
                    if (!response.ok) throw new Error("Upload failed.");
                    
                    const result = await response.json();
                    startPolling(result.run_id);
                } catch (err) {
                    alert("An error occurred during file upload.");
                    submitBtn.disabled = false;
                    submitBtn.textContent = "🚀 Start Sync Task";
                }
            });

            function startPolling(runId) {
                statusBox.style.display = "block";
                downloadContainer.style.display = "none";
                submitBtn.disabled = true;
                submitBtn.textContent = "🔄 Running Sync Engine...";
                
                if (pollInterval) clearInterval(pollInterval);
                
                pollInterval = setInterval(async () => {
                    try {
                        const res = await fetch(`/api/status/${runId}`);
                        if (!res.ok) return;
                        
                        const data = await res.json();
                        
                        // Update Progress
                        const percent = data.progress || 0;
                        progressBar.style.width = `${percent}%`;
                        percentText.textContent = `${percent}%`;
                        rowsCount.textContent = `Processed: ${data.processed_rows || 0}/${data.total_rows || 0} rows`;
                        syncStatus.textContent = `Status: ${data.status.toUpperCase()}`;
                        
                        // Update Logs
                        logBox.innerHTML = (data.logs || []).map(log => `<div>${log}</div>`).join("");
                        logBox.scrollTop = logBox.scrollHeight;
                        
                        if (data.status === "completed") {
                            clearInterval(pollInterval);
                            submitBtn.disabled = false;
                            submitBtn.textContent = "🚀 Start Sync Task";
                            downloadContainer.style.display = "block";
                            downloadBtn.href = `/api/download/${runId}`;
                        } else if (data.status === "failed") {
                            clearInterval(pollInterval);
                            submitBtn.disabled = false;
                            submitBtn.textContent = "🚀 Start Sync Task";
                            alert("Sync runner encountered a fatal error: " + data.error);
                        }
                    } catch (err) {
                        console.error("Polling error:", err);
                    }
                }, 1500);
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content, status_code=200)
