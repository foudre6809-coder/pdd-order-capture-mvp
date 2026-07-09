import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import { Download, FileImage, Loader2, UploadCloud } from "lucide-react";
import "./styles.css";

const API_BASE = "http://127.0.0.1:8000";
const ORDER_FIELDS = ["订单编号", "商品名称", "商品规格", "数量", "订单金额", "订单状态", "下单时间"];

function App() {
  const [files, setFiles] = useState([]);
  const [result, setResult] = useState(null);
  const [status, setStatus] = useState("idle");
  const [error, setError] = useState("");

  const previews = useMemo(
    () =>
      files.map((file) => ({
        name: file.name,
        url: URL.createObjectURL(file),
      })),
    [files],
  );

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

  function downloadExcel() {
    if (!result?.session_id) return;
    window.location.href = `${API_BASE}/api/export/${result.session_id}`;
  }

  const screenshotCount = result?.count ?? files.length;
  const ocrCount = result?.records?.length ?? 0;

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="topbar">
          <div>
            <h1>订单采集工具</h1>
            <p>Phase 1：上传截图、中文 OCR、导出 Excel</p>
          </div>
          <span className={`status ${status}`}>{statusText(status)}</span>
        </header>

        <div className="content-grid">
          <section className="panel upload-panel">
            <div className="panel-title">
              <FileImage size={20} />
              <h2>截图上传</h2>
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

          <section className="panel progress-panel">
            <div className="panel-title">
              <h2>进度</h2>
            </div>
            <ProgressRow label="截图" value={screenshotCount} total={files.length || screenshotCount || 0} />
            <ProgressRow label="OCR" value={ocrCount} total={files.length || ocrCount || 0} />
            <ProgressRow label="解析" value={result ? ocrCount : 0} total={files.length || ocrCount || 0} />

            <button className="secondary-action" type="button" onClick={downloadExcel} disabled={!result?.session_id}>
              <Download size={18} />
              导出 Excel
            </button>
          </section>
        </div>

        <section className="panel results-panel">
          <div className="panel-title">
            <h2>识别结果</h2>
          </div>
          {!result ? (
            <p className="empty-text">上传截图后，这里会显示 OCR 文本和 Excel 字段预填结果。</p>
          ) : (
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
          )}
        </section>
      </section>
    </main>
  );
}

function statusText(status) {
  if (status === "uploading") return "识别中";
  if (status === "done") return "已完成";
  if (status === "error") return "有错误";
  return "待上传";
}

function ProgressRow({ label, value, total }) {
  const denominator = total || 0;
  const percent = denominator ? Math.round((value / denominator) * 100) : 0;
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
