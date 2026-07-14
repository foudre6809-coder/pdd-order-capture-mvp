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
APP_VERSION = "3.0.0"

EXPORT_FIELDS = ["序号", "商品信息", "实付款", "截图文件"]

app = FastAPI(title="拼多多订单截图 OCR 提取工具")
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
    box: list[list[float]] | None = None


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
                OcrLine(text=str(item[1][0]).strip(), score=float(item[1][1]), box=normalize_box(item[0]))
                for item in raw_items
                if item and item[1] and str(item[1][0]).strip()
            ]

        result, _ = engine(str(image_path))
        lines: list[OcrLine] = []
        for item in result or []:
            text = str(item[1]).strip()
            score = float(item[2]) if len(item) > 2 and item[2] is not None else None
            if text:
                lines.append(OcrLine(text=text, score=score, box=normalize_box(item[0])))
        return lines


ocr_engine = OcrEngine()


def ensure_dirs() -> None:
    for path in (SCREENSHOTS_DIR, DATA_DIR, OUTPUTS_DIR):
        path.mkdir(parents=True, exist_ok=True)


def normalize_box(value: Any) -> list[list[float]] | None:
    if not value:
        return None
    try:
        return [[float(point[0]), float(point[1])] for point in value]
    except (TypeError, ValueError, IndexError):
        return None


def box_center_y(line: OcrLine) -> float:
    if not line.box:
        return 0
    return sum(point[1] for point in line.box) / len(line.box)


def box_center_x(line: OcrLine) -> float:
    if not line.box:
        return 0
    return sum(point[0] for point in line.box) / len(line.box)


def line_to_dict(line: OcrLine) -> dict[str, Any]:
    return {"text": line.text, "score": line.score, "box": line.box}


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


def clean_text(value: str) -> str:
    value = value.replace("￥", "¥")
    value = re.sub(r"[ \t]+", " ", value)
    return value.strip(" ：:，,。")


def extract_paid_amount(text: str) -> str:
    normalized = clean_text(text)
    if "实" not in normalized or "付" not in normalized:
        return ""
    match = re.search(r"实\s*付(?:款)?[^0-9¥]{0,8}¥?\s*([0-9]+(?:[.,][0-9]{1,2})?)", normalized)
    if not match:
        match = re.search(r"¥\s*([0-9]+(?:[.,][0-9]{1,2})?)", normalized)
    if not match:
        return ""
    return match.group(1).replace(",", ".")


def chinese_count(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def is_title_candidate(text: str) -> bool:
    value = clean_text(text)
    if chinese_count(value) < 4:
        return False
    reject_words = (
        "我的订单",
        "全部",
        "待付款",
        "拼团中",
        "打包中",
        "待收货",
        "评价",
        "拼单成功",
        "申请退款",
        "再次拼单",
        "催发货",
        "确认收货",
        "查看物流",
        "申请售后",
        "更多",
        "实付",
        "免运费",
        "退货包运费",
        "7天无理由",
        "7天价保",
        "商家正在",
        "预计",
        "快递",
        "旗舰店",
        "百亿补贴",
    )
    if any(word in value for word in reject_words):
        return False
    if re.search(r"(?:¥|￥)\s*\d", value):
        return False
    if re.search(r"^[xX×]\s*\d+$", value):
        return False
    if re.search(r"[|｜]", value):
        return False
    if re.search(r"(黑色|白色|灰色|蓝色|红色|绿色|黄色|紫色|粉色).{0,8}(XL|XXL|L码|M码|S码)", value, re.IGNORECASE):
        return False
    if "【" in value and "】" in value and len(value) < 28:
        return False
    if value.endswith(("店", "店>", "店›", ">")):
        return False
    return True


def title_score(text: str) -> int:
    value = clean_text(text)
    score = chinese_count(value) * 2 + min(len(value), 34)
    if "..." in value or "…" in value:
        score += 10
    if re.search(r"(衣柜|挂衣|吸尘器|猫草|短袖|上衣|手机|家用|车载|零食|盆栽)", value):
        score += 8
    return score


def normalize_title(lines: list[str]) -> str:
    text = "".join(clean_text(line) for line in lines)
    text = re.sub(r"\s+", "", text)
    return text.strip(" ：:，,。")


def group_candidate_lines(lines: list[OcrLine]) -> list[list[OcrLine]]:
    ordered = sorted(lines, key=lambda line: (box_center_y(line), box_center_x(line)))
    groups: list[list[OcrLine]] = []
    for line in ordered:
        if not groups or box_center_y(line) - box_center_y(groups[-1][-1]) > 105:
            groups.append([line])
        else:
            groups[-1].append(line)
    return groups


def extract_visible_price(text: str) -> str:
    match = re.search(r"(?:¥|￥)\s*([0-9]+(?:[.,][0-9]{1,2})?)", text)
    if not match:
        return ""
    return match.group(1).replace(",", ".")


def fallback_visible_price(lines: list[OcrLine], amount_line: OcrLine, lower_bound: float) -> str:
    amount_y = box_center_y(amount_line)
    candidates = [
        line
        for line in lines
        if lower_bound < box_center_y(line) < amount_y and box_center_x(line) > 900 and extract_visible_price(line.text)
    ]
    if not candidates:
        return ""
    price_line = max(candidates, key=box_center_y)
    return extract_visible_price(price_line.text)


def best_title_with_positions(lines: list[OcrLine], amount_line: OcrLine, lower_bound: float) -> str:
    amount_y = box_center_y(amount_line)
    window_top = max(lower_bound, amount_y - 440)
    candidates = [
        line
        for line in lines
        if window_top <= box_center_y(line) < amount_y - 8 and box_center_x(line) > 300 and is_title_candidate(line.text)
    ]
    if not candidates:
        return ""

    best_group = max(
        group_candidate_lines(candidates),
        key=lambda group: (sum(title_score(line.text) for line in group), -abs(box_center_y(group[0]) - amount_y)),
    )
    title_lines = [line.text for line in best_group[:3]]
    return normalize_title(title_lines)


def best_title_without_positions(candidates: list[str]) -> str:
    if not candidates:
        return ""
    return normalize_title([max(candidates[-5:], key=title_score)])


def extract_purchase_items(lines: list[OcrLine], image_name: str, start_index: int = 1) -> list[dict[str, str]]:
    if any(line.box for line in lines):
        ordered = sorted(lines, key=lambda line: (box_center_y(line), box_center_x(line)))
        amount_lines = [line for line in ordered if extract_paid_amount(line.text)]
        items: list[dict[str, str]] = []
        previous_amount_y = 0.0
        for amount_line in amount_lines:
            title = best_title_with_positions(ordered, amount_line, previous_amount_y)
            amount = extract_paid_amount(amount_line.text)
            if "¥" not in clean_text(amount_line.text) and "￥" not in amount_line.text:
                amount = fallback_visible_price(ordered, amount_line, previous_amount_y) or amount
            if title and amount:
                items.append(
                    {
                        "序号": str(start_index + len(items)),
                        "商品信息": title,
                        "实付款": amount,
                        "截图文件": image_name,
                    }
                )
            previous_amount_y = box_center_y(amount_line)
        return items

    items = []
    title_candidates: list[str] = []
    for line in lines:
        amount = extract_paid_amount(line.text)
        if amount:
            title = best_title_without_positions(title_candidates)
            if title:
                items.append(
                    {
                        "序号": str(start_index + len(items)),
                        "商品信息": title,
                        "实付款": amount,
                        "截图文件": image_name,
                    }
                )
            title_candidates = []
        elif is_title_candidate(line.text):
            title_candidates.append(line.text)
    return items


def write_ocr_json(session_id: str, payload: dict[str, Any]) -> Path:
    path = DATA_DIR / f"{session_id}_ocr_result.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def write_excel(session_id: str, items: list[dict[str, str]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单数据"
    sheet.append(EXPORT_FIELDS)
    for item in items:
        sheet.append([item.get(field, "") for field in EXPORT_FIELDS])

    widths = {
        "A": 8,
        "B": 58,
        "C": 14,
        "D": 14,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

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
    all_items: list[dict[str, str]] = []
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
        items = extract_purchase_items(lines, image_name, start_index=len(all_items) + 1)
        all_items.extend(items)
        records.append(
            {
                "image": image_name,
                "ocr_image": ocr_image_path.name,
                "text": text,
                "lines": [line_to_dict(line) for line in lines],
                "items": items,
            }
        )

    payload = {"session_id": session_id, "records": records, "items": all_items}
    ocr_json_path = write_ocr_json(session_id, payload)
    excel_path = write_excel(session_id, all_items)
    return {
        "session_id": session_id,
        "image_count": len(records),
        "item_count": len(all_items),
        "ocr_result": ocr_json_path.name,
        "excel_file": excel_path.name,
        "records": records,
        "items": all_items,
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
