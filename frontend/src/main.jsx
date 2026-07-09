import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import { Download, FileImage, Loader2, Play, RefreshCw, Smartphone, UploadCloud } from "lucide-react";
import "./styles.css";

const API_BASE = "http://127.0.0.1:8000";
const ORDER_FIELDS = ["订单编号", "商品名称", "商品规格", "数量", "订单金额", "订单状态", "下单时间"];

function App() {
  const [files, setFiles] = useState([]);
  const [result, setResult] = useState(null);
  const [status, setStatus] = useState("idle");
  const [error, setError] = useState("");
  const [device, setDevice] = useState(null);
  const [capture, setCapture] = useState(null);
  const [captureError, setCaptureError] = useState("");
  const [screenVersion, setScreenVersion] = useState(Date.now());
  const [settings, setSettings] = useState({
    interval_seconds: 3,
    swipe_distance: 500,
    max_count: 20,
  });

  const previews = useMemo(
    () =>
      files.map((file) => ({
        name: file.name,
        url: URL.createObjectURL(file),
      })),
    [files],
  );

  useEffect(() => {
    refreshDevice();
    refreshCapture();
    const timer = window.setInterval(() => {
      refreshDevice(false);
      refreshCapture();
    }, 3000);
    return () => window.clearInterval(timer);
  }, []);

  async function refreshDevice(updateScreen = true) {
    try {
      const response = await fetch(`${API_BASE}/api/device/status`);
      const payload = await response.json();
      setDevice(payload);
      if (updateScreen && payload.connected) {
        setScreenVersion(Date.now());
      }
    } catch (err) {
      setDevice({ connected: false, message: err.message });
    }
  }

  async function refreshCapture() {
    try {
      const response = await fetch(`${API_BASE}/api/capture/status`);
      const payload = await response.json();
      setCapture(payload);
      if (payload.records?.length) {
        setResult(payload);
      }
    } catch {
      // Progress polling should not interrupt manual upload workflows.
    }
  }

  async function submitFiles() {
    if (!files.length) {
      setError("请先选择订单截图。");
      return;
    }
    setStatus("uploading");
    setError("");
    setResult(null);

    const form = new FormData();
    files.forEach((file) => form.append("files", file));

    try {
      const response = await fetch(`${API_BASE}/api/ocr/upload`, {
        method: "POST",
        body: form,
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.detail || "识别失败");
      }
      setResult(payload);
      setStatus("done");
    } catch (err) {
      setError(err.message);
      setStatus("error");
    }
  }

  async function startCapture() {
    setCaptureError("");
    if (!device?.connected) {
      setCaptureError(device?.message || "手机未连接。");
      return;
    }
    try {
      const response = await fetch(`${API_BASE}/api/capture/start`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(settings),
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.detail || "采集启动失败");
      }
      setCapture(payload);
      setResult(null);
      window.setTimeout(refreshCapture, 500);
    } catch (err) {
      setCaptureError(err.message);
    }
  }

  function downloadExcel() {
    const sessionId = result?.session_id || capture?.session_id;
    if (!sessionId) return;
    window.location.href = `${API_BASE}/api/export/${sessionId}`;
  }

  const captureRunning = capture?.running;
  const screenshotCount = capture?.screenshot_count ?? result?.count ?? files.length;
  const ocrCount = capture?.ocr_count ?? result?.records?.length ?? 0;
  const successCount = capture?.success_order_count ?? countSuccessfulRecords(result?.records);
  const targetCount = capture?.target_count || settings.max_count;
  const deviceConnected = Boolean(device?.connected);

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="topbar">
          <div>
            <h1>订单采集工具</h1>
            <p>Phase 2：手机截图采集、自动滑动、OCR、导出 Excel</p>
          </div>
          <span className={`status ${captureRunning ? "uploading" : status}`}>{captureRunning ? "采集中" : statusText(status)}</span>
        </header>

        <div className="content-grid">
          <section className="panel phone-panel">
            <div className="panel-title">
              <Smartphone size={20} />
              <h2>手机连接</h2>
            </div>

            <div className="device-summary">
              <span className={`dot ${deviceConnected ? "online" : ""}`} />
              <strong>{deviceConnected ? "已连接" : "未连接"}</strong>
              <button className="icon-action" type="button" onClick={() => refreshDevice(true)} aria-label="刷新手机状态">
                <RefreshCw size={16} />
              </button>
            </div>
            <p className="muted-text">
              {deviceConnected
                ? `设备：${device.device_id || "-"}，分辨率：${device.resolution || "未知"}`
                : device?.message || "等待 adb 检测手机。"}
            </p>
            <p className="muted-text">adb：{device?.adb_available ? "可用" : "不可用"} / scrcpy：{device?.scrcpy_available ? "可用" : "不可用"}</p>

            <div className="phone-screen">
              {deviceConnected ? (
                <img src={`${API_BASE}/api/device/screenshot?v=${screenVersion}`} alt="手机画面" />
              ) : (
                <span>手机画面区域</span>
              )}
            </div>
          </section>

          <section className="panel settings-panel">
            <div className="panel-title">
              <h2>采集设置</h2>
            </div>

            <SettingsInput
              label="滑动间隔"
              suffix="秒"
              min={1}
              max={10}
              value={settings.interval_seconds}
              onChange={(value) => setSettings({ ...settings, interval_seconds: value })}
            />
            <SettingsInput
              label="滑动距离"
              suffix="px"
              min={100}
              max={2500}
              value={settings.swipe_distance}
              onChange={(value) => setSettings({ ...settings, swipe_distance: value })}
            />
            <SettingsInput
              label="采集数量"
              suffix="张"
              min={1}
              max={500}
              value={settings.max_count}
              onChange={(value) => setSettings({ ...settings, max_count: value })}
            />

            <button className="primary-action" type="button" onClick={startCapture} disabled={captureRunning}>
              {captureRunning ? <Loader2 className="spin" size={18} /> : <Play size={18} />}
              开始采集
            </button>
            {captureError ? <p className="error-message">{captureError}</p> : null}
          </section>
        </div>

        <section className="panel progress-panel">
          <div className="panel-title">
            <h2>采集进度</h2>
            <span className="step-text">{capture?.current_step || "待开始"}</span>
          </div>
          <div className="progress-grid">
            <ProgressRow label="截图" value={screenshotCount} total={targetCount} />
            <ProgressRow label="OCR" value={ocrCount} total={targetCount} />
            <ProgressRow label="成功订单" value={successCount} total={targetCount} />
            <ProgressRow label="重复跳过" value={capture?.skipped_count || 0} total={targetCount} />
          </div>

          <button className="secondary-action" type="button" onClick={downloadExcel} disabled={!result?.session_id && !capture?.excel_file}>
            <Download size={18} />
            导出 Excel
          </button>
          {capture?.error ? <p className="error-message">{capture.error}</p> : null}
        </section>

        <div className="content-grid">
          <section className="panel upload-panel">
            <div className="panel-title">
              <FileImage size={20} />
              <h2>手动截图上传</h2>
            </div>

            <label className="dropzone">
              <UploadCloud size={28} />
              <span>选择拼多多订单截图</span>
              <input
                type="file"
                accept="image/png,image/jpeg,image/webp"
                multiple
                onChange={(event) => {
                  setFiles(Array.from(event.target.files || []));
                  setResult(null);
                  setError("");
                  setStatus("idle");
                }}
              />
            </label>

            <div className="preview-grid">
              {previews.map((preview) => (
                <figure key={preview.name}>
                  <img src={preview.url} alt={preview.name} />
                  <figcaption>{preview.name}</figcaption>
                </figure>
              ))}
            </div>

            <button className="primary-action" type="button" onClick={submitFiles} disabled={status === "uploading"}>
              {status === "uploading" ? <Loader2 className="spin" size={18} /> : <UploadCloud size={18} />}
              开始 OCR
            </button>

            {error ? <p className="error-message">{error}</p> : null}
          </section>

          <section className="panel results-panel compact-results">
            <div className="panel-title">
              <h2>识别结果</h2>
            </div>
            <Results result={result} />
          </section>
        </div>
      </section>
    </main>
  );
}

function SettingsInput({ label, suffix, min, max, value, onChange }) {
  return (
    <label className="settings-row">
      <span>{label}</span>
      <div>
        <input
          type="number"
          min={min}
          max={max}
          value={value}
          onChange={(event) => onChange(Number(event.target.value))}
        />
        <em>{suffix}</em>
      </div>
    </label>
  );
}

function Results({ result }) {
  if (!result?.records?.length) {
    return <p className="empty-text">采集或上传截图后，这里会显示 OCR 文本和 Excel 字段预填结果。</p>;
  }

  return (
    <div className="result-list">
      {result.records.map((record) => (
        <article className="result-item" key={record.image}>
          <div className="result-header">
            <strong>{record.image}</strong>
            <span>{record.lines.length} 行文本</span>
          </div>
          <div className="field-grid">
            {ORDER_FIELDS.map((field) => (
              <label key={field}>
                <span>{field}</span>
                <input value={record.parsed[field] || ""} readOnly />
              </label>
            ))}
          </div>
          <pre>{record.text || "未识别到文字"}</pre>
        </article>
      ))}
    </div>
  );
}

function countSuccessfulRecords(records = []) {
  return records.filter((record) => record.parsed?.商品名称 && record.parsed?.订单金额).length;
}

function statusText(status) {
  if (status === "uploading") return "识别中";
  if (status === "done") return "已完成";
  if (status === "error") return "有错误";
  return "待上传";
}

function ProgressRow({ label, value, total }) {
  const denominator = total || 0;
  const percent = denominator ? Math.min(100, Math.round((value / denominator) * 100)) : 0;
  return (
    <div className="progress-row">
      <div>
        <span>{label}</span>
        <strong>
          {value} / {denominator}
        </strong>
      </div>
      <div className="bar">
        <span style={{ width: `${percent}%` }} />
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
