from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse
from typing import Optional
import tempfile
import os

app = FastAPI(title="Cleaning Schedule API", version="0.1.0")


@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.post("/generate-cleaning-schedule")
async def generate_cleaning_schedule(
    file: UploadFile = File(...),
    shift_sheet_name: str = Form(...),
    clean_days: str = Form("月水金"),
    holiday_rule: str = Form("non_empty_is_holiday"),
    target_period_label: Optional[str] = Form(None)
):
    """
    Difyから受け取るAPI（まずは受信確認版）
    """
    # 1) 簡単な入力チェック
    if not shift_sheet_name.strip():
        raise HTTPException(status_code=400, detail="shift_sheet_name is required")

    # 2) ファイル拡張子チェック（軽め）
    filename = file.filename or "uploaded_file"
    ext = os.path.splitext(filename)[1].lower()
    if ext not in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
        raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}")

    # 3) 一旦テンポラリ保存（後でExcel解析に使う）
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        # ---- ここから先で後でExcel解析を入れる ----
        # 今は「受け取れたこと」の確認だけ返す

        # 曜日の見やすい表示
        clean_days_normalized = (
            clean_days.replace(",", "")
            .replace("、", "")
            .replace(" ", "")
            .strip()
        )
        if not clean_days_normalized:
            clean_days_normalized = "月水金"

        clean_days_display = "・".join([c for c in clean_days_normalized if c in "月火水木金土日"])

        return JSONResponse({
            "success": True,
            "message": "API受信成功（まだExcel解析は未実装）",
            "received": {
                "filename": filename,
                "shift_sheet_name": shift_sheet_name.strip(),
                "clean_days": clean_days,
                "clean_days_normalized": clean_days_normalized,
                "clean_days_display": clean_days_display,
                "holiday_rule": holiday_rule,
                "target_period_label": target_period_label,
                "saved_temp_path": tmp_path
            }
        })

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Internal error: {str(e)}")