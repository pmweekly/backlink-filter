import React from 'react';
import { createRoot } from 'react-dom/client';
import {
  AlertCircle,
  Check,
  ChevronDown,
  ClipboardList,
  CloudUpload,
  Download,
  FileSpreadsheet,
  History,
  LoaderCircle,
  Play,
  Settings,
  Trash2,
  X,
} from 'lucide-react';
import './styles.css';

const STAGES = [
  { key: 'upload', label: '上传文件' },
  { key: 'processor', label: 'processor 清洗' },
  { key: 'check_blogs', label: '博客评论检测' },
  { key: 'done', label: '导出结果' },
];

const API_BASE = import.meta.env.VITE_API_BASE || '';

function formatBytes(value) {
  if (!Number.isFinite(value)) return '-';
  if (value >= 1024 * 1024) return `${(value / 1024 / 1024).toFixed(1)} MB`;
  if (value >= 1024) return `${(value / 1024).toFixed(1)} KB`;
  return `${value} B`;
}

function formatNumber(value) {
  if (value === undefined || value === null || value === '') return '-';
  return new Intl.NumberFormat('zh-CN').format(value);
}

function stageState(job, stageKey) {
  if (!job) return 'pending';
  if (job.status === 'failed') {
    if (job.stage === stageKey) return 'failed';
    const currentIndex = STAGES.findIndex((stage) => stage.key === job.stage);
    const targetIndex = STAGES.findIndex((stage) => stage.key === stageKey);
    return targetIndex < currentIndex ? 'done' : 'pending';
  }
  if (job.status === 'completed') return 'done';
  const currentIndex = STAGES.findIndex((stage) => stage.key === job.stage);
  const targetIndex = STAGES.findIndex((stage) => stage.key === stageKey);
  if (targetIndex < currentIndex) return 'done';
  if (targetIndex === currentIndex) return 'active';
  return 'pending';
}

function useJobHistory() {
  const [jobs, setJobs] = React.useState([]);

  const loadJobs = React.useCallback(async () => {
    const response = await fetch(`${API_BASE}/api/jobs`);
    if (!response.ok) return;
    setJobs(await response.json());
  }, []);

  React.useEffect(() => {
    loadJobs();
  }, [loadJobs]);

  return { jobs, loadJobs };
}

function App() {
  const fileInputRef = React.useRef(null);
  const eventSourceRef = React.useRef(null);
  const [files, setFiles] = React.useState([]);
  const [job, setJob] = React.useState(null);
  const [logs, setLogs] = React.useState([]);
  const [isDragging, setIsDragging] = React.useState(false);
  const [isUploading, setIsUploading] = React.useState(false);
  const [error, setError] = React.useState('');
  const { jobs, loadJobs } = useJobHistory();

  React.useEffect(() => () => {
    if (eventSourceRef.current) eventSourceRef.current.close();
  }, []);

  const acceptFiles = React.useCallback((nextFiles) => {
    const xlsxFiles = Array.from(nextFiles).filter((file) => file.name.toLowerCase().endsWith('.xlsx'));
    setError(xlsxFiles.length === nextFiles.length ? '' : '仅支持上传 .xlsx 文件');
    setFiles((current) => {
      const byKey = new Map(current.map((file) => [`${file.name}-${file.size}`, file]));
      xlsxFiles.forEach((file) => byKey.set(`${file.name}-${file.size}`, file));
      return Array.from(byKey.values());
    });
  }, []);

  const connectEvents = React.useCallback((jobId) => {
    if (eventSourceRef.current) eventSourceRef.current.close();
    const source = new EventSource(`${API_BASE}/api/jobs/${jobId}/events`);
    eventSourceRef.current = source;
    source.onmessage = (event) => {
      const data = JSON.parse(event.data);
      if (data.type === 'snapshot') {
        setJob(data.payload);
        if (data.payload.status === 'completed' || data.payload.status === 'failed') {
          source.close();
          loadJobs();
        }
      }
      if (data.type === 'log') {
        setLogs((current) => [...current, data.payload].slice(-300));
      }
    };
    source.onerror = () => {
      source.close();
    };
  }, [loadJobs]);

  const runJob = async () => {
    if (!files.length || isUploading) return;
    setIsUploading(true);
    setError('');
    setLogs([]);
    const formData = new FormData();
    files.forEach((file) => formData.append('files', file));
    try {
      const response = await fetch(`${API_BASE}/api/jobs`, {
        method: 'POST',
        body: formData,
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.detail || '上传失败');
      }
      setJob(payload);
      connectEvents(payload.job_id);
      setFiles([]);
    } catch (err) {
      setError(err.message || '上传失败');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const removeFile = (file) => {
    setFiles((current) => current.filter((item) => item !== file));
  };

  const latestCompletedJobs = jobs.filter((item) => item.status === 'completed').slice(0, 5);
  const currentStats = job?.stats?.check_blogs || {};
  const processorStats = job?.stats?.processor || {};
  const progress = job ? Math.round(job.progress || 0) : 0;

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark"><ClipboardList size={22} /></span>
          <span>外链清洗与筛选</span>
        </div>
        <nav className="top-actions" aria-label="顶部操作">
          <button
            className="ghost-button"
            type="button"
            onClick={() => document.getElementById('history-title')?.scrollIntoView({ behavior: 'smooth' })}
          >
            <History size={17} />历史记录
          </button>
          <button
            className="ghost-button"
            type="button"
            onClick={() => document.getElementById('processing-options')?.scrollIntoView({ behavior: 'smooth' })}
          >
            <Settings size={17} />设置
          </button>
          <span className="operator">OP</span>
          <button className="user-button" type="button">运营 <ChevronDown size={15} /></button>
        </nav>
      </header>

      <main className="workspace">
        <section className="panel upload-panel" aria-labelledby="upload-title">
          <div className="panel-header">
            <h1 id="upload-title">上传 XLSX 文件</h1>
          </div>
          <div
            className={`drop-zone ${isDragging ? 'dragging' : ''}`}
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setIsDragging(false);
              acceptFiles(event.dataTransfer.files);
            }}
          >
            <CloudUpload size={40} strokeWidth={1.8} />
            <strong>将 XLSX 文件拖拽到此处，或点击选择文件</strong>
            <span>支持 xlsx 格式，单个文件不超过 100MB</span>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx"
              multiple
              onChange={(event) => acceptFiles(event.target.files || [])}
            />
          </div>

          <div className="toolbar-row">
            <button className="secondary-button" type="button" onClick={() => fileInputRef.current?.click()}>
              <FileSpreadsheet size={16} />选择文件
            </button>
            <button className="secondary-button" type="button" onClick={() => setFiles([])} disabled={!files.length}>
              <Trash2 size={16} />清空列表
            </button>
          </div>

          {error ? (
            <div className="alert" role="alert"><AlertCircle size={17} />{error}</div>
          ) : null}

          <div className="section-title">
            <h2>已选择的文件（{files.length}）</h2>
          </div>
          <div className="file-table" role="table" aria-label="已选择的文件">
            <div className="file-row table-head" role="row">
              <span>文件名</span>
              <span>大小</span>
              <span>状态</span>
              <span>操作</span>
            </div>
            {files.length ? files.map((file) => (
              <div className="file-row" role="row" key={`${file.name}-${file.size}`}>
                <span className="file-name"><FileSpreadsheet size={18} />{file.name}</span>
                <span>{formatBytes(file.size)}</span>
                <span className="ready"><Check size={15} />就绪</span>
                <button className="icon-button" type="button" aria-label={`移除 ${file.name}`} onClick={() => removeFile(file)}>
                  <X size={16} />
                </button>
              </div>
            )) : (
              <div className="empty-row">暂无文件</div>
            )}
          </div>

          <div className="option-panel" id="processing-options">
            <div>
              <span>检测模式</span>
              <strong>单次博客评论检测</strong>
            </div>
            <div>
              <span>输出格式</span>
              <strong>彩色 Excel 工作簿</strong>
            </div>
          </div>

          <button className="primary-button" type="button" onClick={runJob} disabled={!files.length || isUploading}>
            {isUploading ? <LoaderCircle className="spin" size={20} /> : <Play size={20} fill="currentColor" />}
            {isUploading ? '正在上传' : '开始处理'}
          </button>
        </section>

        <section className="right-column">
          <div className="panel progress-panel" aria-labelledby="progress-title">
            <div className="panel-header progress-heading">
              <h2 id="progress-title">处理进度</h2>
              <div className="job-meta">
                <span>任务 ID：{job?.job_id || '-'}</span>
                <span>开始时间：{job?.created_at || '-'}</span>
              </div>
            </div>

            <div className="stage-line" aria-label="处理阶段">
              {STAGES.map((stage, index) => {
                const state = stageState(job, stage.key);
                return (
                  <div className={`stage ${state}`} key={stage.key}>
                    <span className="stage-dot">
                      {state === 'done' ? <Check size={18} /> : state === 'failed' ? <AlertCircle size={17} /> : index + 1}
                    </span>
                    <strong>{stage.label}</strong>
                    <small>{state === 'done' ? '完成' : state === 'active' ? '进行中' : state === 'failed' ? '失败' : '等待中'}</small>
                  </div>
                );
              })}
            </div>

            <div className="progress-block">
              <div className="progress-label">
                <span>总体进度</span>
                <strong>{progress}%</strong>
              </div>
              <div className="bar-track"><span style={{ width: `${progress}%` }} /></div>
            </div>

            <div className="metric-grid">
              <div><span>原始行数</span><strong>{formatNumber(processorStats.rows_read ?? currentStats.raw_rows)}</strong></div>
              <div><span>预过滤垃圾</span><strong>{formatNumber(processorStats.filtered_rows)}</strong></div>
              <div><span>已检测行数</span><strong>{formatNumber(currentStats.processed_rows)}</strong></div>
              <div><span>网络检测</span><strong>{formatNumber(currentStats.network_checked_rows)}</strong></div>
              <div><span>缓存命中</span><strong>{formatNumber(currentStats.cache_hit_rows)}</strong></div>
              <div><span>断点恢复</span><strong>{formatNumber(currentStats.resumed_rows)}</strong></div>
              <div><span>Google 登录</span><strong>{formatNumber(currentStats.google_login_rows)}</strong></div>
              <div><span>博客网站</span><strong>{formatNumber(currentStats.label_counts?.['博客网站'])}</strong></div>
            </div>

            <div className="log-header">
              <h3>实时日志</h3>
              <button className="secondary-button compact" type="button" onClick={() => setLogs([])}>
                <Trash2 size={14} />清空日志
              </button>
            </div>
            <div className="log-console" aria-live="polite">
              {logs.length ? logs.map((entry, index) => (
                <div className={`log-line ${entry.level?.toLowerCase()}`} key={`${entry.time}-${index}`}>
                  <span>{entry.time}</span>
                  <strong>{entry.level}</strong>
                  <p>{entry.message}</p>
                </div>
              )) : (
                <div className="log-placeholder">等待任务开始</div>
              )}
            </div>
          </div>

          <div className="panel history-panel" aria-labelledby="history-title">
            <div className="panel-header">
              <h2 id="history-title">已完成的任务</h2>
            </div>
            <div className="history-table">
              <div className="history-row table-head">
                <span>任务 ID</span>
                <span>开始时间</span>
                <span>输入文件</span>
                <span>检测行数</span>
                <span>状态</span>
                <span>操作</span>
              </div>
              {latestCompletedJobs.length ? latestCompletedJobs.map((item) => (
                <div className="history-row" key={item.job_id}>
                  <span>{item.job_id}</span>
                  <span>{item.created_at}</span>
                  <span>{item.files?.length || 0} 个文件</span>
                  <span>{formatNumber(item.stats?.check_blogs?.processed_rows)}</span>
                  <span className="ready"><Check size={15} />完成</span>
                  <a className="download-button" href={`${API_BASE}/api/jobs/${item.job_id}/download`}>
                    <Download size={16} />下载 Excel
                  </a>
                </div>
              )) : (
                <div className="empty-row">暂无已完成任务</div>
              )}
            </div>
            {job?.status === 'completed' ? (
              <a className="primary-download" href={`${API_BASE}/api/jobs/${job.job_id}/download`}>
                <Download size={18} />下载当前任务 Excel
              </a>
            ) : null}
            {job?.status === 'failed' ? (
              <div className="alert failed" role="alert"><AlertCircle size={17} />{job.error || '任务失败'}</div>
            ) : null}
          </div>
        </section>
      </main>
    </div>
  );
}

createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
