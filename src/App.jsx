import React, { useState, useRef } from 'react';

function Toggle({ label, sublabel, checked, onChange }) {
  return (
    <label className="toggle-row">
      <div className="toggle-text">
        <span className="toggle-label">{label}</span>
        {sublabel && <span className="toggle-sublabel">{sublabel}</span>}
      </div>
      <div className={`toggle-switch ${checked ? 'on' : ''}`} onClick={() => onChange(!checked)}>
        <div className="toggle-knob" />
      </div>
    </label>
  );
}

function RadioGroup({ label, options, value, onChange }) {
  return (
    <div className="input-group">
      <h3 className="section-title">{label}</h3>
      <div className="radio-group">
        {options.map(opt => (
          <label key={opt.value} className="radio-label">
            <input
              type="radio"
              name={label}
              value={opt.value}
              checked={value === opt.value}
              onChange={() => onChange(opt.value)}
            />
            <span>{opt.label}</span>
            {opt.sub && <span className="opt-sub">{opt.sub}</span>}
          </label>
        ))}
      </div>
    </div>
  );
}

export default function App() {
  const [file, setFile]           = useState(null);
  const [genre, setGenre]         = useState('N');
  const [font, setFont]           = useState('Georgia');
  const [spacing, setSpacing]     = useState('1.5');
  const [centerH1, setCenterH1]   = useState(true);
  const [centerH2, setCenterH2]   = useState(false);
  const [borderH1, setBorderH1]   = useState(true);
  const [borderH2, setBorderH2]   = useState(false);
  const [status, setStatus]       = useState('');
  const [isError, setIsError]     = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragOver  = e => { e.preventDefault(); e.currentTarget.classList.add('drag-active'); };
  const handleDragLeave = e => { e.preventDefault(); e.currentTarget.classList.remove('drag-active'); };
  const handleDrop      = e => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-active');
    if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
  };

  const handleFile = f => {
    if (!f.name.endsWith('.docx')) {
      setIsError(true); setStatus('ERROR — ONLY .DOCX FILES ACCEPTED'); return;
    }
    setFile(f); setStatus(''); setIsError(false);
  };

  const handleSubmit = async () => {
    if (!file) { setIsError(true); setStatus('ERROR — NO FILE SELECTED'); return; }
    setIsProcessing(true); setIsError(false); setStatus('PROCESSING MANUSCRIPT...');

    const fd = new FormData();
    fd.append('file', file);
    fd.append('genre', genre);
    fd.append('font', font);
    fd.append('spacing', spacing);
    fd.append('center_h1', centerH1);
    fd.append('center_h2', centerH2);
    fd.append('border_h1', borderH1);
    fd.append('border_h2', borderH2);

    try {
      const res = await fetch('/api/format', { method: 'POST', body: fd });
      if (!res.ok) {
        const txt = await res.text();
        throw new Error(txt);
      }
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = `KDP_Formatted_${file.name}`;
      document.body.appendChild(a); a.click();
      URL.revokeObjectURL(url); a.remove();
      setStatus('DONE — FORMATTED DOCUMENT DOWNLOADED.');
    } catch (err) {
      console.error(err);
      setIsError(true); setStatus('ERROR — PROCESSING FAILED. SEE CONSOLE.');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="app-container">

      {/* ── HEADER ── */}
      <header>
        <div className="header-tag">KDP EBOOK FORMATTER</div>
        <h1>Format Your<br/>Manuscript.</h1>
        <p className="subtitle">Premium document engine. Drop your raw .docx file, configure your layout, and download a professionally formatted KDP-ready eBook in seconds.</p>
      </header>

      {/* ── UPLOAD ── */}
      <section className="upload-zone"
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
      >
        <div className="upload-icon">
          {file ? '✓' : '↑'}
        </div>
        <h2>{file ? file.name : 'DROP YOUR MANUSCRIPT'}</h2>
        <p>{file ? `${(file.size / 1024).toFixed(0)} KB — Click to change` : 'Click to browse or drag & drop (.docx only)'}</p>
        <input type="file" ref={fileInputRef} onChange={e => e.target.files?.[0] && handleFile(e.target.files[0])} accept=".docx" style={{ display: 'none' }} />
      </section>

      {/* ── SETTINGS GRID ── */}
      <section className="settings-grid">

        <RadioGroup
          label="Book Genre"
          value={genre}
          onChange={setGenre}
          options={[
            { value: 'N', label: 'Non-Fiction', sub: 'Block paragraphs, 6pt spacing' },
            { value: 'F', label: 'Fiction',     sub: '0.3″ indent, no gap between paras' },
          ]}
        />

        <RadioGroup
          label="Body Font"
          value={font}
          onChange={setFont}
          options={[
            { value: 'Georgia',          label: 'Georgia 12pt',          sub: 'Premium reference standard' },
            { value: 'Times New Roman',  label: 'Times New Roman 11pt',  sub: 'Classic KDP serif' },
            { value: 'Garamond',         label: 'Garamond 12pt',         sub: 'Elegant editorial' },
          ]}
        />

        <RadioGroup
          label="Line Spacing"
          value={spacing}
          onChange={setSpacing}
          options={[
            { value: '1.5', label: '1.5× (Recommended)', sub: 'Extracted from premium reference' },
            { value: '1.0', label: '1.0× Single',        sub: 'Dense, compact layout' },
            { value: '2.0', label: '2.0× Double',        sub: 'Academic / review copies' },
          ]}
        />

        <div className="input-group">
          <h3 className="section-title">Chapter Headings (H1)</h3>
          <div className="toggle-list">
            <Toggle label="Center H1 Headings"   sublabel="Recommended for KDP eBooks"  checked={centerH1}  onChange={setCenterH1} />
            <Toggle label="Separator Line Below H1" sublabel="Adds a subtle #CCCCCC rule" checked={borderH1}  onChange={setBorderH1} />
          </div>
        </div>

        <div className="input-group">
          <h3 className="section-title">Sub-Headings (H2)</h3>
          <div className="toggle-list">
            <Toggle label="Center H2 Sub-Headings"    sublabel="Left-aligned by default"      checked={centerH2}  onChange={setCenterH2} />
            <Toggle label="Separator Line Below H2"   sublabel="Adds a subtle #CCCCCC rule"   checked={borderH2}  onChange={setBorderH2} />
          </div>
        </div>

      </section>

      {/* ── WHAT'S APPLIED INFO ── */}
      <section className="info-strip">
        <div className="info-item"><span className="info-dot" />Deep-strip manual fonts & colors</div>
        <div className="info-item"><span className="info-dot" />Smart chapter regex detection</div>
        <div className="info-item"><span className="info-dot" />Premium section labels (#888888 Arial)</div>
        <div className="info-item"><span className="info-dot" />Key takeaway bordered boxes</div>
        <div className="info-item"><span className="info-dot" />Auto-updating Table of Contents</div>
        <div className="info-item"><span className="info-dot" />Track changes & comment removal</div>
      </section>

      {/* ── ACTION ── */}
      <button className="action-btn" onClick={handleSubmit} disabled={!file || isProcessing}>
        {isProcessing
          ? <><span className="loader" /> FORMATTING MANUSCRIPT...</>
          : 'FORMAT & DOWNLOAD'}
      </button>

      {status && (
        <div className={`status-message ${isError ? 'status-error' : 'status-ok'}`}>
          {status}
        </div>
      )}

    </div>
  );
}
