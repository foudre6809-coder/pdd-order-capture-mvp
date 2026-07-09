# 订单录屏智能采集与结构化提取工具（Phase 2）

本阶段实现截图上传、OCR 前图片增强、字段提取、Excel 导出，以及基于 adb 的手机截图采集和自动滑动。

不包含视频录制、AI 字段解析、云端部署和用户系统。

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

Phase 2 手机采集需要本机可用的 adb：

```bash
adb devices
```

如果要使用 scrcpy 预览手机画面，可另行安装 scrcpy；本工具不会录制视频，采集流程仍使用 adb 截图。

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

## 输出文件

- 原始截图：`screenshots/<session_id>/001.png`
- OCR 预处理图：`screenshots/<session_id>/ocr_001.png`
- OCR 结果：`work/data/<session_id>_ocr_result.json`
- Excel：`outputs/<session_id>_订单数据.xlsx`

## 接口

- `GET /api/health`：健康检查
- `GET /api/config`：返回当前 OCR 引擎、版本和运行状态
- `GET /api/device/status`：返回 adb/scrcpy 可用性、手机连接状态和分辨率
- `GET /api/device/screenshot`：获取当前手机截图
- `POST /api/capture/start`：启动自动截图采集
- `GET /api/capture/status`：读取采集进度
- `POST /api/capture/stop`：请求停止采集
- `POST /api/ocr/upload`：上传截图并生成 OCR/Excel
- `GET /api/export/<session_id>`：下载 Excel

## 自动采集流程

1. adb 截图
2. 等待页面稳定
3. 保存截图
4. 比较相邻截图差异，重复页面跳过 OCR
5. OCR 和字段提取
6. adb 滑动
7. 按配置循环，连续 3 次无变化会自动结束

## 样例数据

`samples/` 中包含一套本地生成的测试样例：

- `sample_order.png`：示例订单截图
- `ocr_result.json`：OCR 测试结果
- `expected_orders.xlsx`：预期 Excel 结果
