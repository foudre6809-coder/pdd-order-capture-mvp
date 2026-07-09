from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import threading
import time
import uuid
from dataclasses import asdict
from dataclasses import dataclass
from datetime import datetime
from importlib import metadata, util
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import Workbook
from pydantic import BaseModel, Field
from PIL import Image, ImageChops, ImageEnhance, ImageFilter, ImageOps, ImageStat

BASE_DIR = Path(__file__).resolve().parents[1]
SCREENSHOTS_DIR = BASE_DIR / "screenshots"
DATA_DIR = BASE_DIR / "work" / "data"
OUTPUTS_DIR = BASE_DIR / "outputs"
APP_VERSION = "2.0.0"
PAGE_CHANGE_THRESHOLD = 0.015
STABLE_WAIT_SECONDS = 0.5
MAX_UNCHANGED_CAPTURES = 3

ORDER_FIELDS = [
    "订单编号",
    "商品名称",
    "商品规格",
    "数量",
    "订单金额",
    "订单状态",
    "下单时间",
]

app = FastAPI(title="订单录屏智能采集与结构化提取工具 - Phase 2")
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


@dataclass
class DeviceStatus:
    adb_available: bool
    scrcpy_available: bool
    connected: bool
    device_id: str | None = None
    resolution: str | None = None
    message: str = ""


@dataclass
class CaptureJobState:
    session_id: str = ""
    running: bool = False
    status: str = "idle"
    current_step: str = "待开始"
    screenshot_count: int = 0
    ocr_count: int = 0
    success_order_count: int = 0
    skipped_count: int = 0
    target_count: int = 0
    error: str = ""
    excel_file: str = ""
    ocr_result: str = ""
    records: list[dict[str, Any]] | None = None


class CaptureSettings(BaseModel):
    interval_seconds: float = Field(default=3, ge=1, le=10)
    swipe_distance: int = Field(default=500, ge=100, le=2500)
    max_count: int = Field(default=20, ge=1, le=500)


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
capture_state = CaptureJobState(records=[])
capture_lock = threading.Lock()
stop_capture_event = threading.Event()


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


def command_exists(command: str) -> bool:
    return shutil.which(command) is not None


def run_adb(args: list[str], timeout: int = 10, binary: bool = False) -> subprocess.CompletedProcess:
    adb_path = shutil.which("adb")
    if not adb_path:
        raise RuntimeError("未找到 adb，请先安装 Android Platform Tools。")
    return subprocess.run(
        [adb_path, *args],
        check=False,
        capture_output=True,
        timeout=timeout,
        text=not binary,
    )


def get_device_status() -> DeviceStatus:
    adb_available = command_exists("adb")
    scrcpy_available = command_exists("scrcpy")
    if not adb_available:
        return DeviceStatus(
            adb_available=False,
            scrcpy_available=scrcpy_available,
            connected=False,
            message="未找到 adb，请先安装 Android Platform Tools。",
        )

    try:
        result = run_adb(["devices"], timeout=5)
    except Exception as exc:
        return DeviceStatus(adb_available=True, scrcpy_available=scrcpy_available, connected=False, message=str(exc))

    devices: list[tuple[str, str]] = []
    for line in result.stdout.splitlines()[1:]:
        parts = line.split()
        if len(parts) >= 2:
            devices.append((parts[0], parts[1]))

    connected = [device_id for device_id, state in devices if state == "device"]
    if not connected:
        states = ", ".join(f"{device_id}:{state}" for device_id, state in devices)
        return DeviceStatus(
            adb_available=True,
            scrcpy_available=scrcpy_available,
            connected=False,
            message=f"未检测到可用设备。{states}" if states else "未检测到设备。",
        )

    device_id = connected[0]
    resolution = None
    size_result = run_adb(["-s", device_id, "shell", "wm", "size"], timeout=5)
    match = re.search(r"Physical size:\s*(\d+x\d+)", size_result.stdout)
    if match:
        resolution = match.group(1)

    return DeviceStatus(
        adb_available=True,
        scrcpy_available=scrcpy_available,
        connected=True,
        device_id=device_id,
        resolution=resolution,
        message="已连接",
    )


def capture_phone_image(device_id: str) -> Image.Image:
    result = run_adb(["-s", device_id, "exec-out", "screencap", "-p"], timeout=15, binary=True)
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="ignore") if isinstance(result.stderr, bytes) else result.stderr
        raise RuntimeError(stderr.strip() or "手机截图失败。")
    data = result.stdout.replace(b"\r\n", b"\n")
    return Image.open(BytesIO(data)).convert("RGB")


def save_phone_screenshot(device_id: str, path: Path) -> None:
    image = capture_phone_image(device_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path, optimize=True)


def image_difference_ratio(left: Path, right: Path) -> float:
    with Image.open(left) as left_image, Image.open(right) as right_image:
        left_gray = ImageOps.grayscale(left_image).resize((128, 128))
        right_gray = ImageOps.grayscale(right_image).resize((128, 128))
        diff = ImageChops.difference(left_gray, right_gray)
        stat = ImageStat.Stat(diff)
        return (stat.mean[0] or 0) / 255


def swipe_phone(device_id: str, distance: int, resolution: str | None) -> None:
    width, height = 1080, 1920
    if resolution:
        match = re.search(r"(\d+)x(\d+)", resolution)
        if match:
            width, height = int(match.group(1)), int(match.group(2))
    x = width // 2
    start_y = int(height * 0.75)
    end_y = max(int(height * 0.2), start_y - distance)
    result = run_adb(["-s", device_id, "shell", "input", "swipe", str(x), str(start_y), str(x), str(end_y), "350"], timeout=10)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "滑动失败。")


def is_successful_order(parsed: dict[str, str]) -> bool:
    return bool(parsed.get("商品名称") and parsed.get("订单金额"))


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


def set_capture_state(**kwargs: Any) -> None:
    with capture_lock:
        for key, value in kwargs.items():
            setattr(capture_state, key, value)


def get_capture_state_payload() -> dict[str, Any]:
    with capture_lock:
        payload = asdict(capture_state)
        payload["records"] = capture_state.records or []
        return payload


def run_capture_job(settings: CaptureSettings, device: DeviceStatus) -> None:
    assert device.device_id
    session_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8]
    session_dir = SCREENSHOTS_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, Any]] = []
    previous_image: Path | None = None
    unchanged_count = 0

    set_capture_state(
        session_id=session_id,
        running=True,
        status="running",
        current_step="准备采集",
        screenshot_count=0,
        ocr_count=0,
        success_order_count=0,
        skipped_count=0,
        target_count=settings.max_count,
        error="",
        excel_file="",
        ocr_result="",
        records=records,
    )

    try:
        for attempt in range(1, settings.max_count + 1):
            if stop_capture_event.is_set():
                set_capture_state(status="stopped", current_step="已停止")
                break

            image_name = next_screenshot_name(attempt, ".png")
            image_path = session_dir / image_name
            first_path = session_dir / f"stable_check_{attempt:03d}.png"
            ocr_image_path = session_dir / f"ocr_{Path(image_name).stem}.png"

            set_capture_state(current_step="截图")
            save_phone_screenshot(device.device_id, first_path)
            time.sleep(STABLE_WAIT_SECONDS)
            save_phone_screenshot(device.device_id, image_path)
            first_path.unlink(missing_ok=True)

            if previous_image and image_difference_ratio(previous_image, image_path) < PAGE_CHANGE_THRESHOLD:
                unchanged_count += 1
                image_path.unlink(missing_ok=True)
                set_capture_state(skipped_count=capture_state.skipped_count + 1, current_step="页面未变化，跳过 OCR")
            else:
                unchanged_count = 0
                previous_image = image_path
                set_capture_state(screenshot_count=capture_state.screenshot_count + 1, current_step="OCR 处理中")
                preprocess_image(image_path, ocr_image_path)
                lines = ocr_engine.recognize(ocr_image_path)
                text = "\n".join(line.text for line in lines)
                parsed = parse_order_fields(text)
                record = {
                    "image": image_name,
                    "ocr_image": ocr_image_path.name,
                    "text": text,
                    "lines": [{"text": line.text, "score": line.score} for line in lines],
                    "parsed": parsed,
                }
                records.append(record)
                ocr_json_path = write_ocr_json(session_id, records)
                excel_path = write_excel(session_id, records)
                set_capture_state(
                    ocr_count=capture_state.ocr_count + 1,
                    success_order_count=sum(1 for item in records if is_successful_order(item["parsed"])),
                    ocr_result=ocr_json_path.name,
                    excel_file=excel_path.name,
                    records=records,
                    current_step="OCR 完成",
                )

            if unchanged_count >= MAX_UNCHANGED_CAPTURES:
                set_capture_state(status="completed", current_step="页面连续无变化，已结束")
                break

            if attempt < settings.max_count and not stop_capture_event.is_set():
                set_capture_state(current_step="滑动")
                swipe_phone(device.device_id, settings.swipe_distance, device.resolution)
                time.sleep(settings.interval_seconds)
        else:
            set_capture_state(status="completed", current_step="已完成")

        if capture_state.status == "running":
            set_capture_state(status="completed", current_step="已完成")
        if not records:
            ocr_json_path = write_ocr_json(session_id, records)
            excel_path = write_excel(session_id, records)
            set_capture_state(ocr_result=ocr_json_path.name, excel_file=excel_path.name)
    except Exception as exc:
        set_capture_state(running=False, status="error", current_step="采集失败", error=str(exc), records=records)
        return

    set_capture_state(running=False, records=records)


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


@app.get("/api/device/status")
def device_status() -> dict[str, Any]:
    return asdict(get_device_status())


@app.get("/api/device/screenshot")
def device_screenshot() -> FileResponse:
    device = get_device_status()
    if not device.connected or not device.device_id:
        raise HTTPException(status_code=503, detail=device.message or "手机未连接。")
    preview_path = BASE_DIR / "work" / "device_screen.png"
    save_phone_screenshot(device.device_id, preview_path)
    return FileResponse(preview_path, media_type="image/png", filename="device_screen.png")


@app.post("/api/capture/start")
def start_capture(settings: CaptureSettings) -> dict[str, Any]:
    with capture_lock:
        if capture_state.running:
            raise HTTPException(status_code=409, detail="采集任务正在运行。")

    device = get_device_status()
    if not device.connected or not device.device_id:
        raise HTTPException(status_code=503, detail=device.message or "手机未连接。")

    stop_capture_event.clear()
    thread = threading.Thread(target=run_capture_job, args=(settings, device), daemon=True)
    thread.start()
    return get_capture_state_payload()


@app.get("/api/capture/status")
def capture_status() -> dict[str, Any]:
    return get_capture_state_payload()


@app.post("/api/capture/stop")
def stop_capture() -> dict[str, Any]:
    stop_capture_event.set()
    set_capture_state(current_step="正在停止")
    return get_capture_state_payload()


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
