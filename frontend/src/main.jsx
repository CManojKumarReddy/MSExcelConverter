import React, { useState, useEffect } from 'react'
import ReactDOM from 'react-dom/client'
import { GoogleOAuthProvider } from '@react-oauth/google'
import App from './App.jsx'
import './App.css'

function Root() {
  const [clientId, setClientId] = useState(null)

  useEffect(() => {
    fetch('/api/config')
      .then(r => r.json())
      .then(data => setClientId(data.google_client_id || ''))
      .catch(() => setClientId(''))
  }, [])

  if (clientId === null) return null  // wait for config before rendering

  return (
    <GoogleOAuthProvider clientId={clientId}>
      <App />
    </GoogleOAuthProvider>
  )
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>,
)
