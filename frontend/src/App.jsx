import React, { useState, useRef, useEffect, useCallback } from 'react'
import axios from 'axios'

// Relative base: in dev, Vite proxies /api/* to the backend (see vite.config.js);
// in production the FastAPI server serves this built app, so /api/* is same-origin.
const BACKEND_URL = ''

const FILE_ICONS = {
  pdf:  '📄',
  docx: '📝',
  doc:  '📝',
  png:  '🖼️',
  jpg:  '🖼️',
  jpeg: '🖼️',
  csv:  '📊',
  txt:  '📃',
}

const SUPPORTED_TYPES = [
  { ext: 'PDF',  icon: '📄', label: 'PDF Documents' },
  { ext: 'DOCX', icon: '📝', label: 'Word Documents' },
  { ext: 'PNG',  icon: '🖼️', label: 'PNG Images (OCR)' },
  { ext: 'JPG',  icon: '🖼️', label: 'JPG Images (OCR)' },
  { ext: 'CSV',  icon: '📊', label: 'CSV Spreadsheets' },
  { ext: 'TXT',  icon: '📃', label: 'Text Files' },
]

function getFileExt(name) {
  return name.split('.').pop().toLowerCase()
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
}

function WelcomeCard() {
  return (
    <div className="welcome-card">
      <div className="welcome-logo">
        <span className="logo-icon">⚡</span>
        <h1>MSExcelConverter</h1>
      </div>
      <p className="welcome-subtitle">
        Any doc to EXCEL converter — upload a file and I'll turn it into a spreadsheet instantly.
      </p>
      <div className="supported-formats">
        <p className="formats-title">Supported formats</p>
        <div className="format-chips">
          {SUPPORTED_TYPES.map(t => (
            <div className="format-chip" key={t.ext}>
              <span>{t.icon}</span>
              <span>{t.ext}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

function MessageBubble({ msg, passwordInput, setPasswordInput, submitPassword, passwordInputRef, onCancel }) {
  const isUser = msg.role === 'user'
  const isBot  = msg.role === 'bot'

  return (
    <div className={`message-row ${isUser ? 'message-row--user' : 'message-row--bot'}`}>
      {isBot && (
        <div className="avatar avatar--bot">⚡</div>
      )}
      <div className={`bubble ${isUser ? 'bubble--user' : 'bubble--bot'}`}>
        {msg.type === 'file-upload' && (
          <div className="file-preview">
            {msg.thumbnailUrl ? (
              <img className="file-preview__thumb" src={msg.thumbnailUrl} alt={msg.fileName} />
            ) : (
              <span className="file-preview__icon">
                {FILE_ICONS[getFileExt(msg.fileName)] || '📁'}
              </span>
            )}
            <div className="file-preview__info">
              <span className="file-preview__name">{msg.fileName}</span>
              <span className="file-preview__size">{msg.fileSize}</span>
            </div>
          </div>
        )}
        {msg.type === 'converting' && (
          <div className="converting-msg">
            <div className="spinner" />
            <div className="converting-msg__text">
              <span>{msg.text}</span>
              {msg.subtext && <span className="converting-hint">{msg.subtext}</span>}
            </div>
            {msg.cancellable && (
              <button
                className="cancel-btn"
                onClick={() => onCancel?.(msg.jobId)}
                title="Remove this conversion from the queue"
              >
                Cancel
              </button>
            )}
          </div>
        )}
        {msg.type === 'success' && (
          <div className="success-msg">
            <span className="success-icon">✅</span>
            <div className="success-text">
              <p>{msg.text}</p>
              {msg.downloadUrl && (
                <a
                  className="download-btn"
                  href={msg.downloadUrl}
                  download={msg.downloadName}
                  target="_blank"
                  rel="noreferrer"
                >
                  <span>⬇</span> Download Excel File
                </a>
              )}
            </div>
          </div>
        )}
        {msg.type === 'error' && (
          <div className="error-msg">
            <span className="error-icon">⚠️</span>
            <span>{msg.text}</span>
          </div>
        )}
        {msg.type === 'password-prompt' && (
          <div className="password-prompt">
            <span className="password-prompt__text">🔒 {msg.text}</span>
            <div className="password-prompt__row">
              <input
                ref={passwordInputRef}
                type="password"
                className="password-input"
                placeholder="Enter PDF password…"
                value={passwordInput}
                onChange={e => setPasswordInput(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && submitPassword()}
              />
              <button className="password-submit-btn" onClick={submitPassword}>
                Unlock
              </button>
            </div>
          </div>
        )}
        {msg.type === 'text' && <span>{msg.text}</span>}
        <span className="bubble__time">{msg.time}</span>
      </div>
      {isUser && (
        <div className="avatar avatar--user">You</div>
      )}
    </div>
  )
}

function TypingIndicator() {
  return (
    <div className="message-row message-row--bot">
      <div className="avatar avatar--bot">⚡</div>
      <div className="bubble bubble--bot typing-bubble">
        <span className="dot" /><span className="dot" /><span className="dot" />
      </div>
    </div>
  )
}

/* Inline icons — crisp at any size, themeable via currentColor. */
const LinkedInIcon = () => (
  <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor" aria-hidden="true">
    <path d="M20.45 20.45h-3.56v-5.57c0-1.33-.03-3.04-1.85-3.04-1.85 0-2.14 1.45-2.14 2.94v5.67H9.35V9h3.41v1.56h.05c.48-.9 1.64-1.85 3.37-1.85 3.6 0 4.27 2.37 4.27 5.46v6.28zM5.34 7.43a2.06 2.06 0 1 1 0-4.13 2.06 2.06 0 0 1 0 4.13zM7.12 20.45H3.56V9h3.56v11.45zM22.22 0H1.77C.79 0 0 .77 0 1.73v20.54C0 23.22.79 24 1.77 24h20.45c.98 0 1.78-.78 1.78-1.73V1.73C24 .77 23.2 0 22.22 0z"/>
  </svg>
)
const MailIcon = () => (
  <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <rect x="2" y="4" width="20" height="16" rx="2" /><path d="m22 7-10 6L2 7" />
  </svg>
)
const PhoneIcon = () => (
  <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.9.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92z" />
  </svg>
)

function AboutModal({ onClose }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const year = new Date().getFullYear()

  return (
    <div className="about-overlay" onClick={onClose}>
      <div
        className="about-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label="About MSExcelConverter"
      >
        <button className="about-close" onClick={onClose} aria-label="Close">×</button>

        <div className="about-hero">
          <div className="about-logo">⚡</div>
          <h2 className="about-title">MSExcelConverter</h2>
          <p className="about-tagline">
            Turn any document into a clean Excel spreadsheet — PDFs, Word docs,
            images, CSVs and more, in seconds.
          </p>
        </div>

        <div className="about-section">
          <span className="about-section__label">Get in touch</span>

          <a
            className="about-card about-card--link"
            href="https://www.linkedin.com/in/ms-solutions-89b883414/"
            target="_blank"
            rel="noreferrer"
          >
            <span className="about-card__icon about-card__icon--linkedin"><LinkedInIcon /></span>
            <div className="about-card__body">
              <span className="about-card__title">LinkedIn</span>
              <span className="about-card__value">MS Solutions</span>
            </div>
            <span className="about-card__arrow" aria-hidden="true">↗</span>
          </a>

          <a className="about-card about-card--link" href="mailto:msexelconverter@gmail.com">
            <span className="about-card__icon about-card__icon--mail"><MailIcon /></span>
            <div className="about-card__body">
              <span className="about-card__title">Email</span>
              <span className="about-card__value">msexelconverter@gmail.com</span>
            </div>
            <span className="about-card__arrow" aria-hidden="true">↗</span>
          </a>

          <div className="about-card">
            <span className="about-card__icon about-card__icon--phone"><PhoneIcon /></span>
            <div className="about-card__body">
              <span className="about-card__title">Phone</span>
              <a className="about-card__value about-card__value--link" href="tel:+919966170117">+91 99661 70117</a>
              <a className="about-card__value about-card__value--link" href="tel:+918374311097">+91 83743 11097</a>
            </div>
          </div>
        </div>

        <div className="about-footer">© {year} MSExcelConverter · Any doc to EXCEL</div>
      </div>
    </div>
  )
}

export default function App() {
  const [messages, setMessages] = useState([])
  const [isDragging, setIsDragging] = useState(false)
  const [isConverting, setIsConverting] = useState(false)
  const [sheetMode, setSheetMode] = useState('single') // 'single' | 'separate'
  const [mergeImageCols, setMergeImageCols] = useState(false)
  const [adminMode, setAdminMode] = useState(false)      // toggled by Ctrl+M+S
  const [showAbout, setShowAbout] = useState(false)      // About modal visibility
  const [useAzure, setUseAzure] = useState(false)         // admin-only: Azure OCR
  const [cloudEngine, setCloudEngine] = useState(undefined) // undefined=unknown, 'gemini'|'azure'|null
  const [selectedExt, setSelectedExt] = useState('')    // extension of last selected file
  const [pendingFile, setPendingFile] = useState(null)   // file awaiting password
  const [passwordInput, setPasswordInput] = useState('')
  const fileInputRef = useRef(null)
  const messagesEndRef = useRef(null)
  const dropZoneRef = useRef(null)
  const passwordInputRef = useRef(null)
  const activeJobRef = useRef(null)   // { jobId, cancelled } for the in-flight conversion

  const now = () => new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })

  const sleep = (ms) => new Promise(r => setTimeout(r, ms))

  const addMessage = (msg) => {
    setMessages(prev => [...prev, { id: Date.now() + Math.random(), ...msg, time: now() }])
  }

  const replaceLastBotMessage = (msg) => {
    setMessages(prev => {
      const copy = [...prev]
      // Find last bot message and replace it
      for (let i = copy.length - 1; i >= 0; i--) {
        if (copy[i].role === 'bot') {
          copy[i] = { ...copy[i], ...msg, time: now() }
          return copy
        }
      }
      return [...copy, { id: Date.now(), role: 'bot', ...msg, time: now() }]
    })
  }

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  // Admin-mode toggle: hold Ctrl and press M + S together (Ctrl+M+S).
  useEffect(() => {
    const down = new Set()
    const onKeyDown = (e) => {
      const k = e.key.toLowerCase()
      if (k === 'm' || k === 's') down.add(k)
      if (e.ctrlKey && down.has('m') && down.has('s')) {
        e.preventDefault()           // Ctrl+S would otherwise trigger browser-save
        setAdminMode(prev => !prev)
        down.clear()
      }
    }
    const onKeyUp = (e) => { down.delete(e.key.toLowerCase()) }
    const clear = () => down.clear()
    window.addEventListener('keydown', onKeyDown)
    window.addEventListener('keyup', onKeyUp)
    window.addEventListener('blur', clear)
    return () => {
      window.removeEventListener('keydown', onKeyDown)
      window.removeEventListener('keyup', onKeyUp)
      window.removeEventListener('blur', clear)
    }
  }, [])

  // Leaving admin mode disables the Azure option so it can't be sent inadvertently.
  // Entering admin mode checks whether Azure is actually configured on the server.
  useEffect(() => {
    if (!adminMode) {
      if (useAzure) setUseAzure(false)
      return
    }
    axios.get(`${BACKEND_URL}/api/cloud-ocr-status`)
      .then(res => setCloudEngine(res.data?.engine ?? null))
      .catch(() => setCloudEngine(null))
  }, [adminMode]) // eslint-disable-line react-hooks/exhaustive-deps

  const handleFile = useCallback(async (file, mode = sheetMode, password = '') => {
    if (!file) return

    const ext = getFileExt(file.name)
    const allowed = ['pdf', 'docx', 'doc', 'png', 'jpg', 'jpeg', 'csv', 'txt']
    if (!allowed.includes(ext)) {
      addMessage({
        role: 'user',
        type: 'file-upload',
        fileName: file.name,
        fileSize: formatBytes(file.size),
      })
      addMessage({
        role: 'bot',
        type: 'error',
        text: `Sorry, ".${ext}" files are not supported. Please upload a PDF, DOCX, PNG, JPG, CSV, or TXT file.`,
      })
      return
    }

    setIsConverting(true)

    const isImage = ['png', 'jpg', 'jpeg'].includes(ext)

    // User bubble showing the uploaded file (with a thumbnail for images)
    addMessage({
      role: 'user',
      type: 'file-upload',
      fileName: file.name,
      fileSize: formatBytes(file.size),
      thumbnailUrl: isImage ? URL.createObjectURL(file) : null,
    })

    // Bot "converting" bubble
    addMessage({
      role: 'bot',
      type: 'converting',
      text: `Converting ${file.name}...`,
      subtext: isImage ? 'Reading the image — the first AI pass can take a few seconds…' : null,
    })

    try {
      const formData = new FormData()
      formData.append('file', file)
      formData.append('mode', mode)
      if (password) formData.append('password', password)
      formData.append('merge_cols', mergeImageCols ? 'true' : 'false')
      formData.append('use_azure', (adminMode && useAzure) ? 'true' : 'false')

      // Enqueue the conversion — returns immediately with a job id.
      const res = await axios.post(`${BACKEND_URL}/api/convert`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 60000,
      })
      const jobId = res.data.job_id
      activeJobRef.current = { jobId, cancelled: false }

      // Poll for status until the job leaves the queue / finishes.
      const result = await pollJob(jobId, file, isImage)

      if (result?.cancelled) return  // cancel handler already updated the UI

      if (result.status === 'done') {
        const { output_filename, message } = result
        const downloadUrl = `${BACKEND_URL}/api/download/${output_filename}`
        replaceLastBotMessage({
          type: 'success',
          text: message || 'Converted by MSExcelConverter',
          downloadUrl,
          downloadName: output_filename,
        })
      } else if (result.status === 'password_required') {
        // PDF is password-protected — ask the user, then resubmit as a new job.
        setPendingFile(file)
        setPasswordInput('')
        replaceLastBotMessage({
          type: 'password-prompt',
          text: result.password_provided
            ? 'Incorrect password. Please try again:'
            : 'This PDF is password-protected. Please enter the password:',
        })
        setTimeout(() => passwordInputRef.current?.focus(), 100)
      } else if (result.status === 'cancelled') {
        replaceLastBotMessage({ type: 'error', text: 'Conversion was cancelled.' })
      } else {
        replaceLastBotMessage({
          type: 'error',
          text: result.detail || 'Something went wrong during conversion. Please try again.',
        })
      }
    } catch (err) {
      let errText = 'Something went wrong during conversion. Please try again.'
      if (err.response?.data?.detail) {
        errText = err.response.data.detail
      } else if (err.code === 'ECONNREFUSED' || err.code === 'ERR_NETWORK') {
        errText = 'Cannot reach the backend server. Make sure the Python FastAPI server is running on port 8000.'
      }
      replaceLastBotMessage({ type: 'error', text: errText })
    } finally {
      activeJobRef.current = null
      setIsConverting(false)
    }
  }, [sheetMode, mergeImageCols, adminMode, useAzure])

  // Poll a job's status, updating the "converting" bubble (queue position +
  // cancel button while queued). Resolves with the final status payload, or
  // { cancelled: true } if the user cancelled mid-wait.
  const pollJob = useCallback(async (jobId, file, isImage) => {
    while (true) {
      if (activeJobRef.current?.cancelled) return { cancelled: true }

      let data
      try {
        const res = await axios.get(`${BACKEND_URL}/api/convert/${jobId}`, { timeout: 30000 })
        data = res.data
      } catch (err) {
        if (err.response?.status === 404) {
          return { status: 'error', detail: 'This conversion expired. Please try again.' }
        }
        throw err
      }

      if (data.status === 'queued') {
        const pos = data.position || 0
        replaceLastBotMessage({
          type: 'converting',
          text: pos > 0
            ? `Waiting in queue — ${pos} conversion${pos > 1 ? 's' : ''} ahead of you…`
            : 'Next in line — starting shortly…',
          subtext: null,
          cancellable: true,
          jobId,
        })
      } else if (data.status === 'processing') {
        replaceLastBotMessage({
          type: 'converting',
          text: `Converting ${file.name}...`,
          subtext: isImage ? 'Reading the image — the first AI pass can take a few seconds…' : null,
          cancellable: false,
          jobId: null,
        })
      } else {
        return data  // done | error | cancelled | password_required
      }

      await sleep(1000)
    }
  }, [])

  // Cancel a still-queued conversion: tell the server to drop it, stop polling,
  // and update the bubble.
  const cancelQueuedJob = useCallback(async (jobId) => {
    if (activeJobRef.current) activeJobRef.current.cancelled = true
    try {
      await axios.delete(`${BACKEND_URL}/api/convert/${jobId}`, { timeout: 30000 })
    } catch {
      // Ignore — the job may have just started or expired; UI is updated regardless.
    }
    replaceLastBotMessage({ type: 'error', text: 'Conversion cancelled — removed from the queue.' })
    setIsConverting(false)
  }, [])

  // Drag & Drop on the whole window
  useEffect(() => {
    const onDragOver = (e) => { e.preventDefault(); setIsDragging(true) }
    const onDragLeave = (e) => {
      if (!e.relatedTarget || e.relatedTarget === document.body) setIsDragging(false)
    }
    const onDrop = (e) => {
      e.preventDefault()
      setIsDragging(false)
      const file = e.dataTransfer.files[0]
      if (file && !isConverting) {
        setSelectedExt(getFileExt(file.name))
        handleFile(file)
      }
    }
    window.addEventListener('dragover', onDragOver)
    window.addEventListener('dragleave', onDragLeave)
    window.addEventListener('drop', onDrop)
    return () => {
      window.removeEventListener('dragover', onDragOver)
      window.removeEventListener('dragleave', onDragLeave)
      window.removeEventListener('drop', onDrop)
    }
  }, [handleFile, isConverting])

  const onInputChange = (e) => {
    const file = e.target.files[0]
    if (file) {
      setSelectedExt(getFileExt(file.name))
      handleFile(file)
    }
    e.target.value = ''
  }

  const submitPassword = () => {
    if (!pendingFile || !passwordInput.trim()) return
    const file = pendingFile
    const pwd = passwordInput.trim()
    setPendingFile(null)
    setPasswordInput('')
    handleFile(file, sheetMode, pwd)
  }

  const startNewConversion = () => {
    if (isConverting) return
    // Release any image-thumbnail object URLs before clearing.
    messages.forEach(m => { if (m.thumbnailUrl) URL.revokeObjectURL(m.thumbnailUrl) })
    setMessages([])
    setPendingFile(null)
    setPasswordInput('')
  }

  const showWelcome = messages.length === 0

  return (
    <div className="app">
      {/* Header */}
      <header className="header">
        <div className="header__brand">
          <span className="header__icon">⚡</span>
          <div className="header__titles">
            <span className="header__title">MSExcelConverter</span>
            <span className="header__subtitle">Any doc to EXCEL converter</span>
          </div>
        </div>
        <div className="header__badges">
          {messages.length > 0 && (
            <button
              className="header__newbtn"
              onClick={startNewConversion}
              disabled={isConverting}
              title="Clear and start a new conversion"
            >
              + New
            </button>
          )}
          {adminMode && <span className="header__badge header__badge--admin">🔓 ADMIN</span>}
          <button
            className="header__aboutbtn"
            onClick={() => setShowAbout(true)}
            title="About Us"
          >
            About Us
          </button>
          <span className="header__badge">AI Converter</span>
        </div>
      </header>

      {showAbout && <AboutModal onClose={() => setShowAbout(false)} />}

      {/* Drag overlay */}
      {isDragging && (
        <div className="drag-overlay">
          <div className="drag-overlay__inner">
            <span className="drag-overlay__icon">📂</span>
            <p>Drop your file to convert</p>
          </div>
        </div>
      )}

      {/* Messages area */}
      <main className="messages-area">
        <div className="messages-inner">
          {showWelcome && <WelcomeCard />}
          {messages.map(msg => (
            <MessageBubble
              key={msg.id}
              msg={msg}
              passwordInput={passwordInput}
              setPasswordInput={setPasswordInput}
              submitPassword={submitPassword}
              passwordInputRef={passwordInputRef}
              onCancel={cancelQueuedJob}
            />
          ))}
          {isConverting && messages[messages.length - 1]?.type !== 'converting' && (
            <TypingIndicator />
          )}
          <div ref={messagesEndRef} />
        </div>
      </main>

      {/* Options bar — sits above the input bar */}
      <div className="options-bar">
        <div className="options-bar__inner">
          {/* Sheet mode toggle */}
          <div className="options-bar__group">
            <span className="options-bar__label">Output</span>
            <div className="sheet-mode-toggle" title="Choose how tables are arranged in the Excel file">
              <button
                className={`mode-btn ${sheetMode === 'single' ? 'mode-btn--active' : ''}`}
                onClick={() => setSheetMode('single')}
                disabled={isConverting}
              >
                📄 Single Sheet
              </button>
              <button
                className={`mode-btn ${sheetMode === 'separate' ? 'mode-btn--active' : ''}`}
                onClick={() => setSheetMode('separate')}
                disabled={isConverting}
              >
                📑 Separate Sheets
              </button>
            </div>
          </div>

          {/* Merge columns — always enabled, works only for images */}
          <div className="options-bar__group">
            <span className="options-bar__label">Image options</span>
            <label className="merge-cols-label" title="When enabled, repeated column groups in images (e.g. Name | Phone | Name | Phone) are stacked into a single normalized table">
              <input
                type="checkbox"
                checked={mergeImageCols}
                onChange={e => setMergeImageCols(e.target.checked)}
                disabled={isConverting}
              />
              <span>Merge repeating columns</span>
            </label>
            <span className="options-bar__info">ℹ️ Only applied to PNG / JPG files</span>
          </div>

          {/* Admin-only: cloud AI OCR engine (revealed via Ctrl+M+S) */}
          {adminMode && (
            <div className="options-bar__group options-bar__group--admin">
              <span className="options-bar__label">Admin · OCR engine</span>
              <label className="merge-cols-label" title="Use a cloud AI engine (Google Gemini or Azure Document Intelligence) for image conversion — much higher accuracy and native table structure. Falls back to Tesseract if no cloud engine is configured on the server.">
                <input
                  type="checkbox"
                  checked={useAzure}
                  onChange={e => setUseAzure(e.target.checked)}
                  disabled={isConverting}
                />
                <span>AI OCR (cloud)</span>
              </label>
              {cloudEngine === 'gemini' && (
                <span className="options-bar__info azure-status azure-status--ok">✓ Gemini ready</span>
              )}
              {cloudEngine === 'azure' && (
                <span className="options-bar__info azure-status azure-status--ok">✓ Azure ready</span>
              )}
              {cloudEngine === null && (
                <span className="options-bar__info azure-status azure-status--warn">
                  ⚠ Not configured — will use Tesseract (see README)
                </span>
              )}
              {cloudEngine === undefined && (
                <span className="options-bar__info">🔓 Admin mode · PNG / JPG only</span>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Input bar */}
      <footer className="input-bar">
        <div className="input-bar__inner">
          <button
            className={`upload-btn ${isConverting ? 'upload-btn--disabled' : ''}`}
            onClick={() => !isConverting && fileInputRef.current?.click()}
            title="Upload a file"
            disabled={isConverting}
          >
            <span className="upload-btn__icon">📎</span>
            <span className="upload-btn__text">
              {isConverting ? 'Converting...' : 'Upload File'}
            </span>
          </button>

          <div className="input-hint">
            {isConverting
              ? 'Processing your file, please wait…'
              : 'Drag & drop a file anywhere, or click Upload File'}
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.docx,.doc,.png,.jpg,.jpeg,.csv,.txt"
            onChange={onInputChange}
            style={{ display: 'none' }}
          />
        </div>
      </footer>
    </div>
  )
}
