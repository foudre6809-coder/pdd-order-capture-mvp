from __future__ import annotations

import json
import os
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime
from importlib import metadata, util
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import Workbook
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

BASE_DIR = Path(__file__).resolve().parents[1]
SCREENSHOTS_DIR = BASE_DIR / "screenshots"
DATA_DIR = BASE_DIR / "work" / "data"
OUTPUTS_DIR = BASE_DIR / "outputs"
APP_VERSION = "1.5.0"

ORDER_FIELDS = [
    "订单编号",
    "商品名称",
    "商品规格",
    "数量",
    "订单金额",
    "订单状态",
    "下单时间",
]

app = FastAPI(title="订单录屏智能采集与结构化提取工具 - Phase 1.5")
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
        self.requested_name = os.getenv("OCR_ENGINE", "rapidocr").strip().lower()

    def _load(self) -> Any:
        if self._engine is None:
            if self.requested_name == "paddleocr":
                self._engine = self._load_paddleocr()
                self.name = "PaddleOCR"
            else:
                self._engine = self._load_rapidocr()
                self.name = "RapidOCR"
        return self._engine

    def _load_paddleocr(self) -> Any:
        try:
            from paddleocr import PaddleOCR
        except ImportError as exc:
            raise RuntimeError("未安装 PaddleOCR，请安装 backend/requirements-paddle.txt 或改用默认 RapidOCR。") from exc
        return PaddleOCR(use_angle_cls=True, lang="ch")

    def _load_rapidocr(self) -> Any:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as exc:
            raise RuntimeError("未安装 RapidOCR，请先安装 backend/requirements.txt。") from exc
        return RapidOCR()

    def configured_name(self) -> str:
        return "PaddleOCR" if self.requested_name == "paddleocr" else "RapidOCR"

    def version(self) -> str:
        package = "paddleocr" if self.requested_name == "paddleocr" else "rapidocr-onnxruntime"
        try:
            return metadata.version(package)
        except metadata.PackageNotFoundError:
            return "未安装"

    def status(self) -> str:
        module = "paddleocr" if self.requested_name == "paddleocr" else "rapidocr_onnxruntime"
        return "ready" if util.find_spec(module) else "missing"

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


def preprocess_image(src: Path, dest: Path) -> None:
    with Image.open(src) as image:
        image = image.convert("RGB")
        image = image.resize((image.width * 2, image.height * 2), Image.Resampling.LANCZOS)
        image = ImageOps.grayscale(image)
        image = ImageEnhance.Contrast(image).enhance(1.8)
        image = ImageEnhance.Sharpness(image).enhance(1.6)
        image = image.filter(ImageFilter.SHARPEN)
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


def clean_ocr_text(text: str) -> str:
    text = text.replace("￥", "¥")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def text_lines(text: str) -> list[str]:
    return [line.strip(" ：:，,。") for line in clean_ocr_text(text).splitlines() if line.strip()]


def normalize_datetime(value: str) -> str:
    value = value.strip(" ：:，,。")
    value = value.replace("年", "-").replace("月", "-").replace("日", "")
    value = value.replace("/", "-").replace(".", "-")
    value = re.sub(r"(\d{4}-\d{1,2}-\d{1,2})(\d{1,2}:\d{2})", r"\1 \2", value)
    value = re.sub(r"\s+", " ", value)
    match = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:\s*(\d{1,2}):(\d{2})(?::(\d{2}))?)?", value)
    if not match:
        return value
    year, month, day, hour, minute, second = match.groups()
    date = f"{year}-{int(month):02d}-{int(day):02d}"
    if hour and minute:
        return f"{date} {int(hour):02d}:{minute}:{second or '00'}"
    return date


def infer_order_id(lines: list[str], text: str) -> str:
    labeled = extract_first(
        [
            r"订单编号[:：]?\s*([A-Za-z0-9-]{6,})",
            r"订单号[:：]?\s*([A-Za-z0-9-]{6,})",
        ],
        text,
    )
    if labeled:
        return labeled
    for line in lines:
        match = re.search(r"\b([0-9]{12,24})\b", line)
        if match and not re.search(r"20\d{2}[-/.年]", line):
            return match.group(1)
    return ""


def infer_amount(lines: list[str]) -> str:
    priority_words = ("实付", "实付款", "支付", "合计", "订单金额", "付款")
    candidates: list[tuple[int, float, str]] = []
    for line in lines:
        for match in re.finditer(r"(?:¥|￥)?\s*([0-9]+(?:\.[0-9]{1,2})?)", line):
            raw = match.group(1)
            amount = float(raw)
            if amount <= 0:
                continue
            priority = 10 if any(word in line for word in priority_words) else 1
            if "优惠" in line or "减" in line or "券" in line:
                priority -= 3
            candidates.append((priority, amount, raw))
    if not candidates:
        return ""
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return f"{candidates[0][1]:.2f}"


def infer_status(lines: list[str]) -> str:
    statuses = [
        "待付款",
        "待分享",
        "待发货",
        "打包中",
        "待收货",
        "已发货",
        "待评价",
        "已完成",
        "交易成功",
        "拼单成功",
        "退款中",
        "退款成功",
        "已退款",
        "已取消",
        "已关闭",
    ]
    for line in lines:
        for status in statuses:
            if status in line:
                return status
    return ""


def infer_order_time(lines: list[str], text: str) -> str:
    labeled = extract_first(
        [
            r"(?:下单时间|创建时间|订单时间|成交时间|支付时间)[:：]?\s*([0-9]{4}[-/.年][0-9]{1,2}[-/.月][0-9]{1,2}[^\n]*)",
        ],
        text,
    )
    if labeled:
        return normalize_datetime(labeled)
    for line in lines:
        match = re.search(r"([0-9]{4}[-/.年][0-9]{1,2}[-/.月][0-9]{1,2}(?:\s*[0-9]{1,2}:[0-9]{2}(?::[0-9]{2})?)?)", line)
        if match:
            return normalize_datetime(match.group(1))
    return ""


def infer_title(lines: list[str], text: str) -> str:
    labeled = extract_first(
        [
            r"商品(?:名称)?[:：]\s*([^\n]+)",
            r"标题[:：]\s*([^\n]+)",
        ],
        text,
    )
    if labeled:
        return labeled

    reject_words = (
        "拼多多",
        "订单详情",
        "订单编号",
        "订单号",
        "实付",
        "支付",
        "合计",
        "订单金额",
        "下单时间",
        "创建时间",
        "订单时间",
        "规格",
        "颜色",
        "尺码",
        "数量",
        "查看物流",
        "联系商家",
        "联系卖家",
        "申请售后",
        "更多",
    )
    status = infer_status(lines)
    candidates: list[str] = []
    for line in lines:
        if status and status in line:
            continue
        if any(word in line for word in reject_words):
            continue
        if re.search(r"20\d{2}[-/.年]\d{1,2}[-/.月]\d{1,2}", line):
            continue
        if re.fullmatch(r"[A-Za-z0-9 -]{6,}", line):
            continue
        if re.search(r"(?:¥|￥)\s*\d", line):
            continue
        if len(re.findall(r"[\u4e00-\u9fff]", line)) >= 4:
            candidates.append(line)
    if not candidates:
        return ""
    return max(candidates, key=len)


def infer_spec(lines: list[str], text: str) -> str:
    labeled = extract_first(
        [
            r"规格[:：]\s*([^\n]+)",
            r"(?:颜色|尺码|尺寸)[:：]\s*([^\n]+)",
        ],
        text,
    )
    if labeled:
        return labeled
    for line in lines:
        if re.search(r"(黑色|白色|灰色|蓝色|红色|绿色|黄色|紫色|粉色|XL|XXL|L码|M码|S码|均码)", line, re.IGNORECASE):
            if len(line) <= 40 and not re.search(r"(实付|支付|合计|订单|下单|20\d{2})", line):
                return line.strip(" xX×0123456789")
    return ""


def infer_quantity(lines: list[str], text: str) -> str:
    value = extract_first(
        [
            r"[xX×]\s*(\d+)",
            r"数量[:：]?\s*(\d+)",
            r"共\s*(\d+)\s*件",
        ],
        text,
    )
    if value:
        return value
    return "1"


def parse_order_fields(text: str) -> dict[str, str]:
    text = clean_ocr_text(text)
    lines = text_lines(text)
    compact = re.sub(r"[ \t]+", " ", text)
    return {
        "订单编号": infer_order_id(lines, compact),
        "商品名称": infer_title(lines, text),
        "商品规格": infer_spec(lines, text),
        "数量": infer_quantity(lines, compact),
        "订单金额": infer_amount(lines),
        "订单状态": infer_status(lines),
        "下单时间": infer_order_time(lines, text),
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


@app.get("/api/config")
def config() -> dict[str, str | bool]:
    return {
        "app_version": APP_VERSION,
        "ocr_engine": ocr_engine.configured_name(),
        "runtime_ocr_engine": ocr_engine.name if ocr_engine._engine is not None else ocr_engine.configured_name(),
        "loaded": ocr_engine._engine is not None,
        "ocr_engine_version": ocr_engine.version(),
        "status": ocr_engine.status(),
    }


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
        ocr_image_path = session_dir / f"ocr_{Path(image_name).stem}.png"

        with raw_path.open("wb") as buffer:
            shutil.copyfileobj(upload.file, buffer)

        try:
            normalize_image(raw_path, image_path)
            preprocess_image(image_path, ocr_image_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"{upload.filename} 不是可识别的图片。") from exc
        finally:
            raw_path.unlink(missing_ok=True)

        try:
            lines = ocr_engine.recognize(ocr_image_path)
        except RuntimeError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"OCR 识别失败：{exc}") from exc

        text = "\n".join(line.text for line in lines)
        records.append(
            {
                "image": image_name,
                "ocr_image": ocr_image_path.name,
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
