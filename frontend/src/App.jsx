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
        <h1>DocToExcel</h1>
      </div>
      <p className="welcome-subtitle">
        Upload any document and I'll convert it to an Excel spreadsheet instantly.
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
            <span className="file-preview__icon">
              {FILE_ICONS[getFileExt(msg.fileName)] || '📁'}
            </span>
            <div className="file-preview__info">
              <span className="file-preview__name">{msg.fileName}</span>
              <span className="file-preview__size">{msg.fileSize}</span>
            </div>
          </div>
        )}
        {msg.type === 'converting' && (
          <div className="converting-msg">
            <div className="spinner" />
            <span>{msg.text}</span>
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

    // User bubble showing the uploaded file
    addMessage({
      role: 'user',
      type: 'file-upload',
      fileName: file.name,
      fileSize: formatBytes(file.size),
    })

    // Bot "converting" bubble
    addMessage({
      role: 'bot',
      type: 'converting',
      text: `Converting ${file.name}...`,
    })

    try {
      const formData = new FormData()
      formData.append('file', file)
      formData.append('mode', mode)
      if (password) formData.append('password', password)
      formData.append('merge_cols', mergeImageCols ? 'true' : 'false')

      const res = await axios.post(`${BACKEND_URL}/api/convert`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 120000,
      })

      const { output_filename, message } = res.data
      const downloadUrl = `${BACKEND_URL}/api/download/${output_filename}`

      replaceLastBotMessage({
        type: 'success',
        text: message || `Successfully converted "${file.name}" to Excel!`,
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
  }, [sheetMode])

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

  const showWelcome = messages.length === 0

  return (
    <div className="app">
      {/* Header */}
      <header className="header">
        <div className="header__brand">
          <span className="header__icon">⚡</span>
          <span className="header__title">DocToExcel</span>
        </div>
        <span className="header__badge">AI Converter</span>
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
