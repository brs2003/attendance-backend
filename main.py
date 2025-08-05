# from fastapi import FastAPI, UploadFile, File
# from fastapi.responses import JSONResponse
# from fastapi.middleware.cors import CORSMiddleware
# from attendance_logic import process_attendance
# import os

# app = FastAPI()

# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=[
#     "http://localhost:3000",
#     "https://brs2003.github.io",
# ],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# @app.post("/upload")
# async def upload_files(trivandrum: UploadFile = File(...), kochi: UploadFile = File(...)):
#     try:
#         os.makedirs("temp", exist_ok=True)
#         trivandrum_path = f"temp/{trivandrum.filename}"
#         kochi_path = f"temp/{kochi.filename}"

#         with open(trivandrum_path, "wb") as f:
#             f.write(await trivandrum.read())

#         with open(kochi_path, "wb") as f:
#             f.write(await kochi.read())

#         # Pass both file paths for processing
#         result = process_attendance(trivandrum_path, kochi_path)
#         return JSONResponse(result)

#     except Exception as e:
#         print("❌ Error during upload:", e)
#         return JSONResponse(status_code=500, content={"error": str(e)})
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from attendance_logic import process_attendance
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "https://brs2003.github.io",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def upload_files(trivandrum: UploadFile = File(...), kochi: UploadFile = File(...)):
    try:
        os.makedirs("temp", exist_ok=True)
        trivandrum_path = f"temp/{trivandrum.filename}"
        kochi_path = f"temp/{kochi.filename}"

        with open(trivandrum_path, "wb") as f:
            f.write(await trivandrum.read())

        with open(kochi_path, "wb") as f:
            f.write(await kochi.read())

        # Process both files
        result = process_attendance(trivandrum_path, kochi_path)
        return JSONResponse(result)

    except Exception as e:
        print("❌ Error during upload:", e)
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.get("/download")
async def download_report(month: str):
    try:
        folder = os.path.join("backend/data", month)
        file_path = os.path.join(folder, "master.xlsx")

        if os.path.exists(file_path):
            return FileResponse(
                path=file_path,
                filename=f"master_{month}.xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            return JSONResponse(status_code=404, content={"error": f"No report found for {month}"})
    except Exception as e:
        print("❌ Error during download:", e)
        return JSONResponse(status_code=500, content={"error": str(e)})
