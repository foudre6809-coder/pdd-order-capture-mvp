# 订单录屏智能采集与结构化提取工具（Phase 1）

本阶段只实现截图上传、OCR 识别和 Excel 导出，不包含手机连接、自动滚动和 AI 字段解析。

## 运行方式

后端：

```bash
cd backend
/Users/changlifan/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

说明：后端优先使用 PaddleOCR；若本机未安装 PaddleOCR，则使用 RapidOCR ONNX 作为轻量中文 OCR 引擎。

前端：

```bash
cd frontend
pnpm install
pnpm dev
```

访问：

```text
http://127.0.0.1:5173
```

## 输出文件

- 原始截图：`screenshots/<session_id>/001.png`
- OCR 结果：`work/data/<session_id>_ocr_result.json`
- Excel：`outputs/<session_id>_订单数据.xlsx`
