# 拼多多订单截图 OCR 工具

本版本只做三件事：

1. 上传拼多多订单截图
2. OCR 识别并按截图顺序提取“商品信息”和“实付款”
3. 导出 Excel

不包含手机连接、自动滑动、视频录制、AI 字段解析、云端部署和用户系统。

## 运行方式

后端：

```bash
cd backend
/Users/changlifan/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

说明：默认 OCR 引擎是 RapidOCR ONNX，对应 `backend/requirements.txt`。

如需切换 PaddleOCR，可额外安装：

```bash
pip install -r requirements-paddle.txt
OCR_ENGINE=paddleocr uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

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

## 使用方式

1. 在 iPhone 上打开拼多多“我的订单”页面。
2. 截一屏订单截图，或用投屏工具把当前屏幕截成图片。
3. 在网页中上传截图。
4. 点击“开始 OCR”。
5. 检查下方表格，确认商品信息和实付款顺序。
6. 点击“导出 Excel”。

## 输出字段

Excel 只包含：

- 序号
- 商品信息
- 实付款
- 截图文件

## 输出文件

- 原始截图：`screenshots/<session_id>/001.png`
- OCR 预处理图：`screenshots/<session_id>/ocr_001.png`
- OCR 结果：`work/data/<session_id>_ocr_result.json`
- Excel：`outputs/<session_id>_订单数据.xlsx`

## 接口

- `GET /api/health`：健康检查
- `GET /api/config`：返回当前 OCR 引擎、版本和运行状态
- `POST /api/ocr/upload`：上传截图并生成 OCR/Excel
- `GET /api/export/<session_id>`：下载 Excel

## 样例数据

`samples/` 中包含一套本地生成的测试样例：

- `sample_order.png`：示例订单截图
- `ocr_result.json`：OCR 测试结果
- `expected_orders.xlsx`：预期 Excel 结果
