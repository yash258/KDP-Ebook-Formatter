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
            <input type="radio" name={label} value={opt.value}
              checked={value === opt.value} onChange={() => onChange(opt.value)} />
            <span>{opt.label}</span>
            {opt.sub && <span className="opt-sub">{opt.sub}</span>}
          </label>
        ))}
      </div>
    </div>
  );
}

const STAGES = [
  { key: 'clean',    label: 'Cleaning XML'           },
  { key: 'analyze',  label: 'Analyzing Structure'     },
  { key: 'classify', label: 'Classifying Paragraphs'  },
  { key: 'style',    label: 'Applying Premium Styles' },
  { key: 'toc',      label: 'Generating TOC'          },
  { key: 'done',     label: 'Done'                    },
];

function ProcessingStages({ stage }) {
  const current = STAGES.findIndex(s => s.key === stage);
  return (
    <div className="stages">
      {STAGES.map((s, i) => (
        <div key={s.key} className={`stage-item ${i < current ? 'done' : ''} ${i === current ? 'active' : ''}`}>
          <span className="stage-dot">{i < current ? '✓' : i === current ? '◆' : '○'}</span>
          <span className="stage-label">{s.label}</span>
        </div>
      ))}
    </div>
  );
}

export default function App() {
  const [file, setFile]         = useState(null);
  const [genre, setGenre]       = useState('N');
  const [font, setFont]         = useState('Georgia');
  const [spacing, setSpacing]   = useState('1.5');
  const [trimSize, setTrimSize] = useState('5.5x8.5');
  const [centerH1, setCenterH1] = useState(true);
  const [centerH2, setCenterH2] = useState(false);
  const [borderH1, setBorderH1] = useState(true);
  const [borderH2, setBorderH2] = useState(false);
  const [dropCaps, setDropCaps] = useState(false);
  const [pageNums, setPageNums] = useState(true);
  const [autoNum,  setAutoNum]  = useState(false);

  const [stage, setStage]           = useState(null);   // null = idle
  const [status, setStatus]         = useState('');
  const [isError, setIsError]       = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragOver  = e => { e.preventDefault(); e.currentTarget.classList.add('drag-active'); };
  const handleDragLeave = e => { e.preventDefault(); e.currentTarget.classList.remove('drag-active'); };
  const handleDrop      = e => {
    e.preventDefault(); e.currentTarget.classList.remove('drag-active');
    if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
  };
  const handleFile = f => {
    if (!f.name.endsWith('.docx')) {
      setIsError(true); setStatus('ERROR — ONLY .DOCX FILES ACCEPTED'); return;
    }
    setFile(f); setStatus(''); setIsError(false); setStage(null);
  };

  const simulateStages = async () => {
    const delays = [300, 600, 900, 400, 300];
    const keys   = ['clean', 'analyze', 'classify', 'style', 'toc'];
    for (let i = 0; i < keys.length; i++) {
      setStage(keys[i]);
      await new Promise(r => setTimeout(r, delays[i]));
    }
  };

  const handleSubmit = async () => {
    if (!file) { setIsError(true); setStatus('ERROR — NO FILE SELECTED'); return; }
    setIsProcessing(true); setIsError(false); setStatus('');

    const stagesPromise = simulateStages();

    const fd = new FormData();
    fd.append('file',      file);
    fd.append('genre',     genre);
    fd.append('font',      font);
    fd.append('spacing',   spacing);
    fd.append('trim_size', trimSize);
    fd.append('center_h1', centerH1);
    fd.append('center_h2', centerH2);
    fd.append('border_h1', borderH1);
    fd.append('border_h2', borderH2);
    fd.append('drop_caps', dropCaps);
    fd.append('page_nums', pageNums);
    fd.append('auto_num',  autoNum);

    try {
      const res = await fetch('/api/format', { method: 'POST', body: fd });
      await stagesPromise;

      if (!res.ok) throw new Error(await res.text());

      setStage('done');
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = `KDP_Formatted_${file.name}`;
      document.body.appendChild(a); a.click();
      URL.revokeObjectURL(url); a.remove();
      setStatus('FORMATTED DOCUMENT DOWNLOADED.');
    } catch (err) {
      console.error(err);
      setIsError(true); setStage(null);
      setStatus('ERROR — PROCESSING FAILED. CHECK CONSOLE.');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="app-container">

      {/* ── HEADER ── */}
      <header>
        <div className="header-tag">KDP EBOOK FORMATTER · SMART ENGINE</div>
        <h1>From Raw Draft<br/>to Premium eBook.</h1>
        <p className="subtitle">
          The engine reads your manuscript, intelligently classifies every paragraph —
          titles, chapters, section labels, key takeaways, block quotes, dividers —
          and applies exact premium formatting so readers love every page.
        </p>
      </header>

      {/* ── UPLOAD ── */}
      <section className="upload-zone"
        onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}>
        <div className="upload-icon">{file ? '✓' : '↑'}</div>
        <h2>{file ? file.name : 'DROP YOUR MANUSCRIPT'}</h2>
        <p>{file ? `${(file.size / 1024).toFixed(0)} KB — Click to replace` : 'Click or drag & drop (.docx only)'}</p>
        <input type="file" ref={fileInputRef} onChange={e => e.target.files?.[0] && handleFile(e.target.files[0])} accept=".docx" style={{ display: 'none' }} />
      </section>

      {/* ── SETTINGS ── */}
      <section className="settings-grid">
        <RadioGroup label="Book Genre" value={genre} onChange={setGenre} options={[
          { value: 'N', label: 'Non-Fiction', sub: 'Block paragraphs · 6pt spacing' },
          { value: 'F', label: 'Fiction',     sub: '0.3″ indent · no gap between paras' },
        ]} />

        <RadioGroup label="Body Font" value={font} onChange={setFont} options={[
          { value: 'Georgia',         label: 'Georgia 12pt',         sub: 'Premium reference standard' },
          { value: 'Times New Roman', label: 'Times New Roman 11pt', sub: 'Classic KDP serif' },
          { value: 'Garamond',        label: 'Garamond 12pt',        sub: 'Elegant editorial' },
        ]} />

        <RadioGroup label="Line Spacing" value={spacing} onChange={setSpacing} options={[
          { value: '1.5', label: '1.5× Spacing', sub: 'Extracted from premium reference' },
          { value: '1.0', label: '1.0× Single',  sub: 'Dense compact layout' },
          { value: '2.0', label: '2.0× Double',  sub: 'Academic / review copies' },
        ]} />

        <div className="input-group">
          <h3 className="section-title">Chapter Headings (H1)</h3>
          <div className="toggle-list">
            <Toggle label="Center Headings" sublabel="Recommended for eBooks" checked={centerH1} onChange={setCenterH1} />
            <Toggle label="Line Below Heading" sublabel="Subtle #CCCCCC rule" checked={borderH1} onChange={setBorderH1} />
          </div>
        </div>

        <div className="input-group">
          <h3 className="section-title">Sub-Headings (H2)</h3>
          <div className="toggle-list">
            <Toggle label="Center Sub-Headings" sublabel="Left-aligned by default" checked={centerH2} onChange={setCenterH2} />
            <Toggle label="Line Below Sub-Heading" sublabel="Subtle #CCCCCC rule" checked={borderH2} onChange={setBorderH2} />
          </div>
        </div>

        <RadioGroup label="KDP Trim Size" value={trimSize} onChange={setTrimSize} options={[
          { value: '5x8',     label: '5″ × 8″',       sub: 'Novels & self-help' },
          { value: '5.5x8.5', label: '5.5″ × 8.5″',   sub: 'Non-fiction standard' },
          { value: '6x9',     label: '6″ × 9″',       sub: 'Business & textbooks' },
        ]} />

        <div className="input-group">
          <h3 className="section-title">Publishing Options</h3>
          <div className="toggle-list">
            <Toggle label="Page Numbers" sublabel="Centered in footer" checked={pageNums} onChange={setPageNums} />
            <Toggle label="Drop Caps" sublabel="First letter of each chapter" checked={dropCaps} onChange={setDropCaps} />
            <Toggle label="Auto-Number Chapters" sublabel="Adds Chapter 1, 2, 3…" checked={autoNum} onChange={setAutoNum} />
          </div>
        </div>

      </section>

      {/* ── WHAT THE ENGINE DETECTS ── */}
      <section className="detect-strip">
        <div className="detect-title">WHAT THE ENGINE DETECTS</div>
        <div className="detect-grid">
          {[
            ['TITLE',        'Book title on first page'],
            ['SUBTITLE',     'Secondary title line'],
            ['CHAPTER H1',   'Chapter / Part headings'],
            ['SECTION H2',   'Bold short sub-sections'],
            ['LABEL',        'Section labels (WHY IT WORKS)'],
            ['KEY TAKEAWAY', 'Bordered quote boxes'],
            ['BLOCK QUOTE',  'Indented italic passages'],
            ['DIVIDER',      'Scene breaks → border lines'],
            ['BODY',         'Justified paragraph text'],
            ['CLEANUP',      'Tracks changes & comments'],
          ].map(([tag, desc]) => (
            <div key={tag} className="detect-item">
              <span className="detect-tag">{tag}</span>
              <span className="detect-desc">{desc}</span>
            </div>
          ))}
        </div>
      </section>

      {/* ── PROCESSING STAGES ── */}
      {stage && <ProcessingStages stage={stage} />}

      {/* ── ACTION ── */}
      <button className="action-btn" onClick={handleSubmit} disabled={!file || isProcessing}>
        {isProcessing
          ? <><span className="loader" /> ANALYZING & FORMATTING...</>
          : 'ANALYZE & FORMAT MANUSCRIPT'}
      </button>

      {status && (
        <div className={`status-message ${isError ? 'status-error' : 'status-ok'}`}>
          {status}
        </div>
      )}

    </div>
  );
}
