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

// Maps each UI section name to the category string stored in Supabase.
// Adjust the values here if your `categories` column uses different spellings.
const SECTION_TO_CATEGORY = {
  'Entrance':       'entrance',
  'Lord Have Mercy':'lord have mercy',
  'Gloria':         'gloria',
  'Acclamation':    'acclamation',
  'Offertory':      'offertory',
  'Holy Holy':      'holy holy',
  'Proclamation':   'proclamation',
  'Communion':      'communion',
  'Recessional':    'recessional',
}

/**
 * Returns the hymns that belong to a given Mass section.
 * Comparison is case-insensitive so "Lord have Mercy" matches "Lord Have Mercy".
 */
const filterHymnsForSection = (hymns, sectionName) => {
  const target = SECTION_TO_CATEGORY[sectionName] ?? sectionName.toLowerCase()
  return hymns.filter(h => (h.categories ?? '').toLowerCase() === target)
}

const TEMPLATE_URL =
  import.meta.env.VITE_TEMPLATE_URL ||
  'https://YOUR_PROJECT.supabase.co/storage/v1/object/public/templates/master.pptx'

const API_URL = import.meta.env.VITE_API_URL || '/api/generate-ppt'

// ---------------------------------------------------------------------------
// Icons (inline SVG)
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
 * Renders one Mass-part row: a hymn dropdown + dynamic verse checkboxes.
 *
 * Props:
 *   index          – 0-based section index
 *   name           – section name string
 *   hymns          – full hymns array already fetched from Supabase
 *   hymnId         – currently selected hymn id (string | '')
 *   selectedVerses – number[] of selected verse numbers
 *   onChange       – (sectionName, hymnId, selectedVerses) => void
 */
function SectionSelector({ index, name, hymns, hymnId, selectedVerses, onChange }) {
  // Only show hymns that belong to this Mass section
  const filteredHymns = filterHymnsForSection(hymns, name)

  // Find the full hymn object from the already-fetched list — no extra DB call needed
  const selectedHymn = hymns.find(h => String(h.id) === String(hymnId)) ?? null

  const verseCount = selectedHymn?.verse_count ?? 0
  const hasChorus  = !!(selectedHymn?.chorus?.trim())

  const handleHymnChange = (e) => {
    const newId = e.target.value
    if (!newId) {
      onChange(name, '', [])
      return
    }
    const hymn = hymns.find(h => String(h.id) === String(newId))
    // Default: all verses selected
    const allVerses = Array.from({ length: hymn?.verse_count ?? 0 }, (_, i) => i + 1)
    onChange(name, newId, allVerses)
  }

  const handleVerseToggle = (verseNum) => {
    const next = selectedVerses.includes(verseNum)
      ? selectedVerses.filter(v => v !== verseNum)
      : [...selectedVerses, verseNum].sort((a, b) => a - b)
    onChange(name, hymnId, next)
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
            value={hymnId}
            onChange={handleHymnChange}
          >
            <option value="">— Select a hymn —</option>
            {filteredHymns.length === 0
              ? <option disabled value="">No hymns in this category yet</option>
              : filteredHymns.map(h => (
                  <option key={h.id} value={h.id}>{h.name}</option>
                ))
            }
          </select>

          {/* Verse checkboxes — rendered directly from fetched data */}
          {hymnId && verseCount > 0 && (
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
                <span className="chorus-badge">♪ Chorus included</span>
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
  const [hymns, setHymns]             = useState([])
  const [hymnsLoading, setHymnsLoading] = useState(true)
  const [hymnsError, setHymnsError]   = useState(null)

  const [date, setDate] = useState(() => {
    // Default to next Sunday
    const today = new Date()
    const day   = today.getDay()           // 0 = Sun
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
  const [statusMsg, setStatusMsg]     = useState(null) // { type: 'error'|'info', text }

  // --- Fetch all hymns on mount ---
  // Selects every column we need so SectionSelector never has to make extra queries.
  useEffect(() => {
    supabase
      .from('hymns')
      .select('id, name, categories, verse_count, chorus, verse_1, verse_2, verse_3, verse_4, verse_5')
      .order('name', { ascending: true })
      .then(({ data, error }) => {
        setHymnsLoading(false)
        if (error) {
          console.error('Supabase error:', error)
          setHymnsError(`Could not load hymns: ${error.message}`)
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

  // --- Build the sections payload for the backend ---
  const buildSectionsPayload = () => {
    return MASS_SECTIONS.map(sectionName => {
      const { hymnId, verses } = selections[sectionName]
      const hymn = hymns.find(h => String(h.id) === String(hymnId))

      if (!hymn) return { name: sectionName, song: { title: '', lyrics: [] } }

      // Build ordered lyrics from flat verse columns + chorus column
      const selectedLyrics = []

      for (const verseNum of verses) {
        const text = hymn[`verse_${verseNum}`]
        if (text?.trim()) {
          selectedLyrics.push({ label: `Verse ${verseNum}`, text: text.trim() })
        }
      }

      // Append chorus if it exists (backend appends it after each verse automatically)
      if (hymn.chorus?.trim()) {
        selectedLyrics.push({ label: 'Chorus', text: hymn.chorus.trim() })
      }

      return {
        name: sectionName,
        song: {
          title:  hymn.name,
          lyrics: selectedLyrics,
        },
      }
    })
  }

  // --- Format date for PPT cover page ---
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

  // --- Validation: enable button if at least one section has a hymn selected ---
  const canDownload = !downloading && !!date &&
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
          ? Array.from({ length: 5 }).map((_, i) => (
              <div key={i} className="skeleton" style={{ marginBottom: '1rem' }} />
            ))
          : MASS_SECTIONS.map((name, index) => (
              <SectionSelector
                key={name}
                index={index}
                name={name}
                hymns={hymns}
                hymnId={selections[name].hymnId}
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
