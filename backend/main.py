from __future__ import annotations

import json
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import Workbook
from PIL import Image

BASE_DIR = Path(__file__).resolve().parents[1]
SCREENSHOTS_DIR = BASE_DIR / "screenshots"
DATA_DIR = BASE_DIR / "work" / "data"
OUTPUTS_DIR = BASE_DIR / "outputs"

ORDER_FIELDS = [
    "订单编号",
    "商品名称",
    "商品规格",
    "数量",
    "订单金额",
    "订单状态",
    "下单时间",
]

app = FastAPI(title="订单录屏智能采集与结构化提取工具 - Phase 1")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@dataclass
class OcrLine:
    text: str
    score: float | None = None


class OcrEngine:
    def __init__(self) -> None:
        self._engine: Any | None = None
        self.name = "未加载"

    def _load(self) -> Any:
        if self._engine is None:
            try:
                from paddleocr import PaddleOCR

                self._engine = PaddleOCR(use_angle_cls=True, lang="ch")
                self.name = "PaddleOCR"
            except ImportError:
                try:
                    from rapidocr_onnxruntime import RapidOCR
                except ImportError as exc:
                    raise RuntimeError("未安装 OCR 引擎，请先安装 PaddleOCR 或 RapidOCR 依赖。") from exc
                self._engine = RapidOCR()
                self.name = "RapidOCR"
        return self._engine

    def recognize(self, image_path: Path) -> list[OcrLine]:
        engine = self._load()
        if self.name == "PaddleOCR":
            result = engine.ocr(str(image_path), cls=True)
            raw_items = result[0] if result else []
            return [
                OcrLine(text=str(item[1][0]).strip(), score=float(item[1][1]))
                for item in raw_items
                if item and item[1] and str(item[1][0]).strip()
            ]

        result, _ = engine(str(image_path))
        lines: list[OcrLine] = []
        for item in result or []:
            text = str(item[1]).strip()
            score = float(item[2]) if len(item) > 2 and item[2] is not None else None
            if text:
                lines.append(OcrLine(text=text, score=score))
        return lines


ocr_engine = OcrEngine()


def ensure_dirs() -> None:
    for path in (SCREENSHOTS_DIR, DATA_DIR, OUTPUTS_DIR):
        path.mkdir(parents=True, exist_ok=True)


def normalize_image(src: Path, dest: Path) -> None:
    with Image.open(src) as image:
        image = image.convert("RGB")
        image.save(dest, optimize=True)


def next_screenshot_name(index: int, suffix: str) -> str:
    ext = suffix.lower()
    if ext not in {".png", ".jpg", ".jpeg", ".webp"}:
        ext = ".png"
    return f"{index:03d}{ext}"


def extract_first(patterns: list[str], text: str) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip(" ：:，,。")
    return ""


def parse_order_fields(text: str) -> dict[str, str]:
    compact = re.sub(r"[ \t]+", " ", text)
    return {
        "订单编号": extract_first(
            [
                r"订单编号[:：]?\s*([A-Za-z0-9-]{6,})",
                r"订单号[:：]?\s*([A-Za-z0-9-]{6,})",
            ],
            compact,
        ),
        "商品名称": extract_first(
            [
                r"商品(?:名称)?[:：]\s*([^\n]+)",
                r"标题[:：]\s*([^\n]+)",
            ],
            text,
        ),
        "商品规格": extract_first(
            [
                r"规格[:：]\s*([^\n]+)",
                r"(?:颜色|尺码|尺寸)[:：]\s*([^\n]+)",
            ],
            text,
        ),
        "数量": extract_first(
            [
                r"[xX×]\s*(\d+)",
                r"数量[:：]?\s*(\d+)",
                r"共\s*(\d+)\s*件",
            ],
            compact,
        ),
        "订单金额": extract_first(
            [
                r"实付款[:：]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
                r"支付[:：]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
                r"[¥￥]\s*([0-9]+(?:\.[0-9]{1,2})?)",
            ],
            compact,
        ),
        "订单状态": extract_first(
            [
                r"(待付款|待发货|待收货|已完成|交易成功|退款成功|已取消|已关闭)",
            ],
            compact,
        ),
        "下单时间": extract_first(
            [
                r"下单时间[:：]?\s*([0-9]{4}[-/.年][0-9]{1,2}[-/.月][0-9]{1,2}[^\n]*)",
                r"([0-9]{4}[-/.年][0-9]{1,2}[-/.月][0-9]{1,2}\s+[0-9]{1,2}:[0-9]{2}(?::[0-9]{2})?)",
            ],
            text,
        ),
    }


def write_ocr_json(session_id: str, records: list[dict[str, Any]]) -> Path:
    path = DATA_DIR / f"{session_id}_ocr_result.json"
    path.write_text(
        json.dumps(records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return path


def write_excel(session_id: str, records: list[dict[str, Any]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单数据"
    headers = [*ORDER_FIELDS, "截图文件", "OCR全文"]
    sheet.append(headers)
    for record in records:
        parsed = record["parsed"]
        sheet.append([*(parsed.get(field, "") for field in ORDER_FIELDS), record["image"], record["text"]])

    for column in sheet.columns:
        max_length = max(len(str(cell.value or "")) for cell in column)
        sheet.column_dimensions[column[0].column_letter].width = min(max(max_length + 2, 12), 60)

    path = OUTPUTS_DIR / f"{session_id}_订单数据.xlsx"
    workbook.save(path)
    return path


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok", "ocr_engine": ocr_engine.name}


@app.post("/api/ocr/upload")
async def upload_screenshots(files: list[UploadFile] = File(...)) -> dict[str, Any]:
    ensure_dirs()
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一张截图。")

    session_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8]
    session_dir = SCREENSHOTS_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    records: list[dict[str, Any]] = []
    for index, upload in enumerate(files, start=1):
        suffix = Path(upload.filename or "").suffix
        image_name = next_screenshot_name(index, suffix)
        raw_path = session_dir / f"raw_{image_name}"
        image_path = session_dir / image_name

        with raw_path.open("wb") as buffer:
            shutil.copyfileobj(upload.file, buffer)

        try:
            normalize_image(raw_path, image_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"{upload.filename} 不是可识别的图片。") from exc
        finally:
            raw_path.unlink(missing_ok=True)

        try:
            lines = ocr_engine.recognize(image_path)
        except RuntimeError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"OCR 识别失败：{exc}") from exc

        text = "\n".join(line.text for line in lines)
        records.append(
            {
                "image": image_name,
                "text": text,
                "lines": [{"text": line.text, "score": line.score} for line in lines],
                "parsed": parse_order_fields(text),
            }
        )

    ocr_json_path = write_ocr_json(session_id, records)
    excel_path = write_excel(session_id, records)
    return {
        "session_id": session_id,
        "count": len(records),
        "ocr_result": ocr_json_path.name,
        "excel_file": excel_path.name,
        "records": records,
    }


@app.get("/api/export/{session_id}")
def export_excel(session_id: str) -> FileResponse:
    matches = sorted(OUTPUTS_DIR.glob(f"{session_id}_订单数据.xlsx"))
    if not matches:
        raise HTTPException(status_code=404, detail="未找到 Excel 文件，请先上传截图识别。")
    return FileResponse(
        matches[0],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="订单数据.xlsx",
    )
