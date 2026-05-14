import React, { useState, useRef } from 'react';

function App() {
  const [file, setFile] = useState(null);
  const [genre, setGenre] = useState('F');
  const [font, setFont] = useState('Georgia');
  const [spacing, setSpacing] = useState('1.5');
  const [status, setStatus] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
    e.currentTarget.classList.add('drag-active');
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-active');
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-active');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  };

  const handleFile = (selectedFile) => {
    if (!selectedFile.name.endsWith('.docx')) {
      setStatus('ERROR: ONLY .DOCX FILES ACCEPTED');
      return;
    }
    setFile(selectedFile);
    setStatus('');
  };

  const handleSubmit = async () => {
    if (!file) {
      setStatus('ERROR: NO FILE SELECTED');
      return;
    }

    setIsProcessing(true);
    setStatus('PROCESSING MANUSCRIPT...');

    const formData = new FormData();
    formData.append('file', file);
    formData.append('genre', genre);
    formData.append('font', font);
    formData.append('spacing', spacing);

    try {
      // Calls the serverless function hosted on Vercel
      const response = await fetch('/api/format', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error(`Server error: ${response.statusText}`);
      }

      // Download the processed file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Formatted_${file.name}`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      a.remove();
      
      setStatus('FORMATTING COMPLETE. DOWNLOAD STARTED.');
    } catch (err) {
      console.error(err);
      setStatus('ERROR: PROCESSING FAILED.');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="app-container">
      <header>
        <h1>KDP Formatter</h1>
        <p className="subtitle">Production-Ready Manuscript Standardizer</p>
      </header>

      <main>
        <div 
          className="upload-zone"
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <h2>{file ? file.name : 'DROP MANUSCRIPT HERE'}</h2>
          <p>{file ? 'CLICK TO CHANGE FILE' : 'OR CLICK TO BROWSE (.DOCX ONLY)'}</p>
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleFileChange} 
            accept=".docx" 
            style={{ display: 'none' }} 
          />
        </div>

        <div className="form-grid">
          <div className="input-group">
            <h3 className="section-title">Genre Rules</h3>
            <div className="radio-group">
              <label className="radio-label">
                <input type="radio" name="genre" value="F" checked={genre === 'F'} onChange={(e) => setGenre(e.target.value)} />
                Fiction (Indents, 0pt spacing)
              </label>
              <label className="radio-label">
                <input type="radio" name="genre" value="N" checked={genre === 'N'} onChange={(e) => setGenre(e.target.value)} />
                Non-Fiction (Block, 6pt spacing)
              </label>
            </div>
          </div>

          <div className="input-group">
            <h3 className="section-title">Typography & Spacing</h3>
            <div className="radio-group">
              <label className="radio-label">
                <input type="radio" name="font" value="Georgia" checked={font === 'Georgia'} onChange={(e) => setFont(e.target.value)} />
                Georgia (Premium Extract)
              </label>
              <label className="radio-label">
                <input type="radio" name="font" value="Times New Roman" checked={font === 'Times New Roman'} onChange={(e) => setFont(e.target.value)} />
                Times New Roman
              </label>
              <label className="radio-label">
                <input type="radio" name="spacing" value="1.5" checked={spacing === '1.5'} onChange={(e) => setSpacing(e.target.value)} />
                1.5 Line Spacing (Default)
              </label>
            </div>
          </div>
        </div>

        <button 
          className="action-btn" 
          onClick={handleSubmit} 
          disabled={!file || isProcessing}
        >
          {isProcessing ? (
            <><span className="loader"></span> FORMATTING...</>
          ) : (
            'FORMAT MANUSCRIPT'
          )}
        </button>

        {status && (
          <div className="status-message">
            {status}
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
