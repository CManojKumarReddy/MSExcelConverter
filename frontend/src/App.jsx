import React, { useState, useRef, useEffect, useCallback } from 'react'
import axios from 'axios'

const BACKEND_URL = 'http://localhost:8000'

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

function MessageBubble({ msg, passwordInput, setPasswordInput, submitPassword, passwordInputRef }) {
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

export default function App() {
  const [messages, setMessages] = useState([])
  const [isDragging, setIsDragging] = useState(false)
  const [isConverting, setIsConverting] = useState(false)
  const [sheetMode, setSheetMode] = useState('single') // 'single' | 'separate'
  const [mergeImageCols, setMergeImageCols] = useState(false)
  const [adminMode, setAdminMode] = useState(false)      // toggled by Ctrl+M+S
  const [useAzure, setUseAzure] = useState(false)         // admin-only: Azure OCR
  const [cloudEngine, setCloudEngine] = useState(undefined) // undefined=unknown, 'gemini'|'azure'|null
  const [selectedExt, setSelectedExt] = useState('')    // extension of last selected file
  const [pendingFile, setPendingFile] = useState(null)   // file awaiting password
  const [passwordInput, setPasswordInput] = useState('')
  const fileInputRef = useRef(null)
  const messagesEndRef = useRef(null)
  const dropZoneRef = useRef(null)
  const passwordInputRef = useRef(null)

  const now = () => new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })

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

      const res = await axios.post(`${BACKEND_URL}/api/convert`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 180000,  // headroom for cloud-OCR retry/backoff waits
      })

      const { output_filename, message } = res.data
      const downloadUrl = `${BACKEND_URL}/api/download/${output_filename}`

      replaceLastBotMessage({
        type: 'success',
        text: message || 'Converted by MSExcelConverter',
        downloadUrl,
        downloadName: output_filename,
      })
    } catch (err) {
      if (err.response?.status === 423) {
        // PDF is password-protected — ask the user
        setPendingFile(file)
        setPasswordInput('')
        replaceLastBotMessage({
          type: 'password-prompt',
          text: password
            ? 'Incorrect password. Please try again:'
            : 'This PDF is password-protected. Please enter the password:',
        })
        setTimeout(() => passwordInputRef.current?.focus(), 100)
      } else {
        let errText = 'Something went wrong during conversion. Please try again.'
        if (err.response?.data?.detail) {
          errText = err.response.data.detail
        } else if (err.code === 'ECONNREFUSED' || err.code === 'ERR_NETWORK') {
          errText = 'Cannot reach the backend server. Make sure the Python FastAPI server is running on port 8000.'
        }
        replaceLastBotMessage({ type: 'error', text: errText })
      }
    } finally {
      setIsConverting(false)
    }
  }, [sheetMode, mergeImageCols, adminMode, useAzure])

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
          <span className="header__badge">AI Converter</span>
        </div>
      </header>

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
