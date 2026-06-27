import { useState, useEffect, useCallback } from 'react'
import { supabase } from './supabaseClient'
import './App.css'

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const MASS_SECTIONS = [
  'Entrance',
  'Lord Have Mercy',
  'Gloria',
  'Acclamation',
  'Offertory',
  'Holy Holy',
  'Proclamation',
  'Communion',
  'Recessional',
]

// The public URL of your master PPTX template stored in Supabase Storage (or any CDN).
// Replace this with your actual template URL.
const TEMPLATE_URL =
  import.meta.env.VITE_TEMPLATE_URL ||
  'https://YOUR_PROJECT.supabase.co/storage/v1/object/public/templates/master.pptx'

// Backend API endpoint (proxied via Vite in dev; absolute URL in production).
const API_URL = import.meta.env.VITE_API_URL || '/api/generate-ppt'

// ---------------------------------------------------------------------------
// Icons (inline SVG – no extra dependency)
// ---------------------------------------------------------------------------

const IconPresentation = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="3" width="20" height="14" rx="2" ry="2" />
    <line x1="8" y1="21" x2="16" y2="21" />
    <line x1="12" y1="17" x2="12" y2="21" />
  </svg>
)

const IconCalendar = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
    <line x1="16" y1="2" x2="16" y2="6" />
    <line x1="8" y1="2" x2="8" y2="6" />
    <line x1="3" y1="10" x2="21" y2="10" />
  </svg>
)

const IconMusic = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M9 18V5l12-2v13" />
    <circle cx="6" cy="18" r="3" />
    <circle cx="18" cy="16" r="3" />
  </svg>
)

const IconDownload = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
)

const IconAlert = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10" />
    <line x1="12" y1="8" x2="12" y2="12" />
    <line x1="12" y1="16" x2="12.01" y2="16" />
  </svg>
)

const IconInfo = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10" />
    <line x1="12" y1="16" x2="12" y2="12" />
    <line x1="12" y1="8" x2="12.01" y2="8" />
  </svg>
)

// ---------------------------------------------------------------------------
// SectionSelector sub-component
// ---------------------------------------------------------------------------

/**
 * Renders one Mass-part row: dropdown + dynamic verse checkboxes.
 *
 * Props:
 *   index      – 0-based section index
 *   name       – section name string
 *   hymns      – full hymns array from Supabase
 *   value      – currently selected hymn id (string | '')
 *   onChange   – (sectionName, hymnId, selectedVerses) => void
 */
function SectionSelector({ index, name, hymns, value, selectedVerses, onChange }) {
  const [verseCount, setVerseCount] = useState(0)
  const [hasChorus, setHasChorus] = useState(false)
  const [loadingVerses, setLoadingVerses] = useState(false)

  // When hymn changes, fetch verse_count + chorus info from Supabase
  useEffect(() => {
    if (!value) {
      setVerseCount(0)
      setHasChorus(false)
      return
    }

    let cancelled = false
    setLoadingVerses(true)

    supabase
      .from('hymns')
      .select('verse_count, has_chorus')
      .eq('id', value)
      .single()
      .then(({ data, error }) => {
        if (cancelled) return
        setLoadingVerses(false)
        if (error || !data) {
          setVerseCount(0)
          setHasChorus(false)
          return
        }
        setVerseCount(data.verse_count ?? 0)
        setHasChorus(data.has_chorus ?? false)
        // Default: all verses selected
        const allVerses = Array.from({ length: data.verse_count ?? 0 }, (_, i) => i + 1)
        onChange(name, value, allVerses)
      })

    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value])

  const handleHymnChange = (e) => {
    const hymnId = e.target.value
    onChange(name, hymnId, [])   // verses reset; effect above will re-populate
  }

  const handleVerseToggle = (verseNum) => {
    const next = selectedVerses.includes(verseNum)
      ? selectedVerses.filter(v => v !== verseNum)
      : [...selectedVerses, verseNum].sort((a, b) => a - b)
    onChange(name, value, next)
  }

  return (
    <div>
      {index > 0 && <div className="section-divider" />}
      <div className="section-row">
        {/* Label */}
        <div className="section-label">
          <span className="section-number">{index + 1}</span>
          <span className="section-name">{name}</span>
        </div>

        {/* Controls */}
        <div className="section-controls">
          <select
            id={`section-${index}`}
            className="styled-select"
            value={value}
            onChange={handleHymnChange}
          >
            <option value="">— Select a hymn —</option>
            {hymns.map(h => (
              <option key={h.id} value={h.id}>{h.title}</option>
            ))}
          </select>

          {/* Verse checkboxes */}
          {loadingVerses && <div className="skeleton" />}

          {!loadingVerses && value && verseCount > 0 && (
            <div className="verse-boxes">
              {Array.from({ length: verseCount }, (_, i) => i + 1).map(num => (
                <label
                  key={num}
                  className={`verse-chip ${selectedVerses.includes(num) ? 'checked' : ''}`}
                >
                  <input
                    type="checkbox"
                    checked={selectedVerses.includes(num)}
                    onChange={() => handleVerseToggle(num)}
                  />
                  <span>Verse {num}</span>
                </label>
              ))}
              {hasChorus && (
                <span className="chorus-badge">
                  ♪ Chorus included
                </span>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Main App
// ---------------------------------------------------------------------------

export default function App() {
  // --- State ---
  const [hymns, setHymns]           = useState([])
  const [hymnsLoading, setHymnsLoading] = useState(true)
  const [hymnsError, setHymnsError]  = useState(null)

  const [date, setDate] = useState(() => {
    // Default to next Sunday
    const today = new Date()
    const day   = today.getDay()               // 0 = Sun
    const diff  = day === 0 ? 0 : 7 - day
    const next  = new Date(today)
    next.setDate(today.getDate() + diff)
    return next.toISOString().split('T')[0]
  })

  // selections[sectionName] = { hymnId: string, verses: number[] }
  const [selections, setSelections] = useState(() =>
    Object.fromEntries(MASS_SECTIONS.map(s => [s, { hymnId: '', verses: [] }]))
  )

  const [downloading, setDownloading] = useState(false)
  const [statusMsg, setStatusMsg]     = useState(null) // { type: 'error'|'info', text: string }

  // --- Fetch hymns on mount ---
  useEffect(() => {
    supabase
      .from('hymns')
      .select('id, title, verse_count, has_chorus, lyrics')
      .order('title', { ascending: true })
      .then(({ data, error }) => {
        setHymnsLoading(false)
        if (error) {
          setHymnsError('Could not load hymns from database. Check your Supabase credentials.')
          return
        }
        setHymns(data ?? [])
      })
  }, [])

  // --- Selection change handler ---
  const handleSelectionChange = useCallback((sectionName, hymnId, verses) => {
    setSelections(prev => ({
      ...prev,
      [sectionName]: { hymnId, verses },
    }))
  }, [])

  // --- Build sections payload ---
  const buildSectionsPayload = () => {
    return MASS_SECTIONS.map(sectionName => {
      const { hymnId, verses } = selections[sectionName]
      const hymn = hymns.find(h => String(h.id) === String(hymnId))

      if (!hymn) return { name: sectionName, song: { title: '', lyrics: [] } }

      // Filter and order the lyrics based on selected verses + chorus
      const rawLyrics = hymn.lyrics ?? []   // expected: [{ label, text }, ...]

      // Separate chorus entries from verses
      const chorusEntry = rawLyrics.find(l => l.label?.toLowerCase().includes('chorus'))

      // Build ordered lyrics: selected verses interleaved with chorus placement
      const selectedLyrics = []
      for (const verseNum of verses) {
        const verseEntry = rawLyrics.find(l => {
          const label = l.label?.toLowerCase() ?? ''
          return label.includes(`verse ${verseNum}`) || label === `v${verseNum}`
        })
        if (verseEntry) selectedLyrics.push(verseEntry)
      }
      if (chorusEntry) selectedLyrics.push(chorusEntry)

      return {
        name: sectionName,
        song: {
          title:  hymn.title,
          lyrics: selectedLyrics,
        },
      }
    })
  }

  // --- Format date for display in PPT (e.g. "Sunday, 29 June 2026") ---
  const formatDate = (isoDate) => {
    const d = new Date(isoDate + 'T00:00:00')
    return d.toLocaleDateString('en-GB', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric',
    })
  }

  // --- Download handler ---
  const handleDownload = async () => {
    setStatusMsg(null)
    setDownloading(true)

    const payload = {
      template_url: TEMPLATE_URL,
      date:         formatDate(date),
      sections:     buildSectionsPayload(),
    }

    try {
      const response = await fetch(API_URL, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(payload),
      })

      if (!response.ok) {
        const err = await response.json().catch(() => ({ error: response.statusText }))
        throw new Error(err.error ?? `Server error ${response.status}`)
      }

      // Trigger browser download
      const blob = await response.blob()
      const url  = URL.createObjectURL(blob)
      const a    = document.createElement('a')
      a.href     = url
      a.download = 'Mass_Presentation.pptx'
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)

      setStatusMsg({ type: 'info', text: 'Presentation downloaded successfully! 🎉' })
    } catch (err) {
      setStatusMsg({ type: 'error', text: `Download failed: ${err.message}` })
    } finally {
      setDownloading(false)
    }
  }

  // --- Validation ---
  const canDownload = !downloading && date &&
    MASS_SECTIONS.some(s => selections[s].hymnId !== '')

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  return (
    <div className="app-wrapper">
      {/* Header */}
      <header className="app-header">
        <div className="logo-icon">
          <IconPresentation />
        </div>
        <h1>Mass Presentation Generator</h1>
        <p>Select hymns for each part of the Mass and generate a ready-to-project PowerPoint.</p>
      </header>

      {/* Status banner */}
      {statusMsg && (
        <div className={`status-banner ${statusMsg.type}`}>
          {statusMsg.type === 'error' ? <IconAlert /> : <IconInfo />}
          {statusMsg.text}
        </div>
      )}

      {/* Hymns fetch error */}
      {hymnsError && (
        <div className="status-banner error">
          <IconAlert />
          {hymnsError}
        </div>
      )}

      {/* ---- Date Card ---- */}
      <div className="card">
        <p className="card-title"><IconCalendar /> Mass Date</p>
        <div className="date-field">
          <label htmlFor="mass-date">Select the Sunday date for this Mass</label>
          <input
            id="mass-date"
            type="date"
            value={date}
            onChange={e => setDate(e.target.value)}
          />
        </div>
      </div>

      {/* ---- Hymn Selections Card ---- */}
      <div className="card">
        <p className="card-title"><IconMusic /> Hymn Selections</p>

        {hymnsLoading
          ? Array.from({ length: 3 }).map((_, i) => <div key={i} className="skeleton" style={{ marginBottom: '1rem' }} />)
          : MASS_SECTIONS.map((name, index) => (
              <SectionSelector
                key={name}
                index={index}
                name={name}
                hymns={hymns}
                value={selections[name].hymnId}
                selectedVerses={selections[name].verses}
                onChange={handleSelectionChange}
              />
            ))
        }
      </div>

      {/* ---- Download Button ---- */}
      <button
        id="download-btn"
        className="download-btn"
        disabled={!canDownload}
        onClick={handleDownload}
      >
        {downloading
          ? <><div className="btn-spinner" /> Generating Presentation…</>
          : <><IconDownload /> Download Mass_Presentation.pptx</>
        }
      </button>

      <footer className="app-footer">
        Built for Sunday liturgy planning &nbsp;·&nbsp; Powered by Supabase + python-pptx
      </footer>
    </div>
  )
}
