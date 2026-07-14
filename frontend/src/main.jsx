import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import { Download, FileImage, Loader2, UploadCloud } from "lucide-react";
import "./styles.css";

const API_BASE = "http://127.0.0.1:8000";

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
      setError("请先选择拼多多订单截图。");
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

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="topbar">
          <div>
            <h1>订单截图 OCR 工具</h1>
            <p>上传一屏拼多多订单截图，按顺序提取商品信息和实付款。</p>
          </div>
          <span className={`status ${status}`}>{statusText(status)}</span>
        </header>

        <div className="content-grid">
          <section className="panel upload-panel">
            <div className="panel-title">
              <FileImage size={20} />
              <h2>上传截图</h2>
            </div>

            <label className="dropzone">
              <UploadCloud size={30} />
              <span>选择或拖入订单截图</span>
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

          <section className="panel summary-panel">
            <div className="panel-title">
              <h2>识别进度</h2>
            </div>
            <div className="metric-grid">
              <Metric label="截图" value={result?.image_count ?? files.length} />
              <Metric label="商品" value={result?.item_count ?? 0} />
              <Metric label="OCR" value={statusText(status)} />
            </div>
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
          <Results result={result} />
        </section>
      </section>
    </main>
  );
}

function Metric({ label, value }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function Results({ result }) {
  const items = result?.items || [];
  if (!items.length) {
    return <p className="empty-text">上传截图并完成 OCR 后，这里会按截图中的上下顺序显示商品信息和实付款。</p>;
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>序号</th>
            <th>商品信息</th>
            <th>实付款</th>
            <th>截图文件</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item) => (
            <tr key={`${item.截图文件}-${item.序号}`}>
              <td>{item.序号}</td>
              <td>{item.商品信息}</td>
              <td>¥{item.实付款}</td>
              <td>{item.截图文件}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function statusText(status) {
  if (status === "uploading") return "识别中";
  if (status === "done") return "已完成";
  if (status === "error") return "有错误";
  return "待上传";
}

createRoot(document.getElementById("root")).render(<App />);
