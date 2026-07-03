import { useState, useEffect, useCallback } from 'react'
import { supabase } from './supabaseClient'
import './App.css'

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const MASS_SECTIONS = [
  'Entrance 1',
  'Entrance 2',
  'Lord Have Mercy',
  'Gloria',
  'Acclamation',
  'Offertory 1',
  'Offertory 2',
  'Holy Holy',
  'Proclamation',
  'Communion 1',
  'Communion 2',
  'Communion 3',
  'Communion 4',
  'Recessional 1',
  'Recessional 2',
]

const SECTION_TO_CATEGORY = {
  'Entrance 1':     'entrance',
  'Entrance 2':     'entrance',
  'Lord Have Mercy':'lord have mercy',
  'Gloria':         'gloria',
  'Acclamation':    'acclamation',
  'Offertory 1':    'offertory',
  'Offertory 2':    'offertory',
  'Holy Holy':      'holy holy',
  'Proclamation':   'proclamation',
  'Communion 1':    'communion',
  'Communion 2':    'communion',
  'Communion 3':    'communion',
  'Communion 4':    'communion',
  'Recessional 1':  'recessional',
  'Recessional 2':  'recessional',
}

const filterHymnsForSection = (hymns, sectionName) => {
  const target = SECTION_TO_CATEGORY[sectionName] ?? sectionName.toLowerCase()
  return hymns.filter(h => (h.categories ?? '').toLowerCase() === target)
}

const TEMPLATE_URL =
  import.meta.env.VITE_TEMPLATE_URL ||
  'https://wvmxlnwfjtesbppojstu.supabase.co/storage/v1/object/public/templates/standard_template.pptx'

// Always use the standard Vercel api/ folder path — never rely on an env var for this.
const API_URL = '/api/generate-ppt'

// ---------------------------------------------------------------------------
// Inline SVG icons
// ---------------------------------------------------------------------------

const ChevronIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="6 9 12 15 18 9" />
  </svg>
)

const CalendarIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2" />
    <line x1="16" y1="2" x2="16" y2="6" />
    <line x1="8" y1="2" x2="8" y2="6" />
    <line x1="3" y1="10" x2="21" y2="10" />
  </svg>
)

const MusicIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d="M9 18V5l12-2v13" />
    <circle cx="6" cy="18" r="3" />
    <circle cx="18" cy="16" r="3" />
  </svg>
)

const DownloadIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
)

const AlertIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10" />
    <line x1="12" y1="8" x2="12" y2="12" />
    <line x1="12" y1="16" x2="12.01" y2="16" />
  </svg>
)

const CheckIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10" />
    <polyline points="9 12 11 14 15 10" />
  </svg>
)

// ---------------------------------------------------------------------------
// SectionSelector — one row per Mass part
// ---------------------------------------------------------------------------

function SectionSelector({ index, name, hymns, hymnId, selectedVerses, onChange }) {
  const filteredHymns = filterHymnsForSection(hymns, name)
  const selectedHymn  = hymns.find(h => String(h.id) === String(hymnId)) ?? null
  const verseCount    = selectedHymn?.verse_count ?? 0
  const hasChorus     = !!(selectedHymn?.chorus?.trim())

  const handleHymnChange = (e) => {
    const newId = e.target.value
    if (!newId) { onChange(name, '', []); return }
    const hymn      = hymns.find(h => String(h.id) === String(newId))
    const allVerses = Array.from({ length: hymn?.verse_count ?? 0 }, (_, i) => i + 1)
    onChange(name, newId, allVerses)
  }

  const handleVerseToggle = (verseNum) => {
    const next = selectedVerses.includes(verseNum)
      ? selectedVerses.filter(v => v !== verseNum)
      : [...selectedVerses, verseNum].sort((a, b) => a - b)
    onChange(name, hymnId, next)
  }

  const isOptional = name.endsWith('2') || name.endsWith('3') || name.endsWith('4')

  return (
    <div className={`section-row ${hymnId ? 'section-row--active' : ''}`}>
      {/* Index badge + name */}
      <div className="section-meta">
        <span className="section-index">{String(index + 1).padStart(2, '0')}</span>
        <div style={{ display: 'flex', flexDirection: 'column', marginTop: '-2px' }}>
          <span className="section-name">{name}</span>
          {isOptional && <span style={{ fontSize: '10px', color: 'var(--text-muted)', lineHeight: '1.2' }}>(Optional)</span>}
        </div>
      </div>

      {/* Dropdown + verse chips */}
      <div className="section-right">
        <div className="select-wrapper">
          <select
            id={`section-${index}`}
            className="styled-select"
            value={hymnId}
            onChange={handleHymnChange}
          >
            <option value="">— Select hymn —</option>
            {filteredHymns.length === 0
              ? <option disabled value="">No hymns in this category</option>
              : filteredHymns.map(h => (
                  <option key={h.id} value={h.id}>{h.name}</option>
                ))
            }
          </select>
          <span className="select-chevron"><ChevronIcon /></span>
        </div>

        {/* Verse checkboxes */}
        {hymnId && verseCount > 0 && (
          <div className="verse-module">
            <div className="verse-chips">
              {Array.from({ length: verseCount }, (_, i) => i + 1).map(num => (
                <label
                  key={num}
                  className={`verse-chip ${selectedVerses.includes(num) ? 'verse-chip--on' : ''}`}
                >
                  <input
                    type="checkbox"
                    checked={selectedVerses.includes(num)}
                    onChange={() => handleVerseToggle(num)}
                  />
                  <span>v{num}</span>
                </label>
              ))}
              {hasChorus && (
                <span className="chorus-pill">♪ chorus</span>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Main App
// ---------------------------------------------------------------------------

export default function App() {
  const [hymns,        setHymns]        = useState([])
  const [hymnsLoading, setHymnsLoading] = useState(true)
  const [hymnsError,   setHymnsError]   = useState(null)

  const [date, setDate] = useState(() => {
    const today = new Date()
    const diff  = today.getDay() === 0 ? 0 : 7 - today.getDay()
    const next  = new Date(today)
    next.setDate(today.getDate() + diff)
    return next.toISOString().split('T')[0]
  })

  const [selections, setSelections] = useState(() =>
    Object.fromEntries(MASS_SECTIONS.map(s => [s, { hymnId: '', verses: [] }]))
  )

  const [downloading, setDownloading] = useState(false)
  const [statusMsg,   setStatusMsg]   = useState(null)

  // Fetch all hymns on mount
  useEffect(() => {
    supabase
      .from('hymns')
      .select('id, name, categories, verse_count, chorus, verse_1, verse_2, verse_3, verse_4, verse_5, chord_link')
      .order('name', { ascending: true })
      .then(({ data, error }) => {
        setHymnsLoading(false)
        if (error) { setHymnsError(`Could not load hymns: ${error.message}`); return }
        setHymns(data ?? [])
      })
  }, [])

  const handleSelectionChange = useCallback((sectionName, hymnId, verses) => {
    setSelections(prev => ({ ...prev, [sectionName]: { hymnId, verses } }))
  }, [])

  const isSectionVisible = (sectionName) => {
    if (sectionName === 'Entrance 2')    return !!selections['Entrance 1']?.hymnId
    if (sectionName === 'Offertory 2')   return !!selections['Offertory 1']?.hymnId
    if (sectionName === 'Communion 4')   return !!selections['Communion 3']?.hymnId
    if (sectionName === 'Recessional 2') return !!selections['Recessional 1']?.hymnId
    return true
  }

  const visibleSections = MASS_SECTIONS.filter(isSectionVisible)

  const buildSectionsPayload = () =>
    MASS_SECTIONS.map(sectionName => {
      const { hymnId, verses } = selections[sectionName]
      const hymn = hymns.find(h => String(h.id) === String(hymnId))
      if (!hymn) return { name: sectionName, song: { title: '', lyrics: [] } }

      const selectedLyrics = []
      for (const verseNum of verses) {
        const text = hymn[`verse_${verseNum}`]
        if (text?.trim()) selectedLyrics.push({ label: `Verse ${verseNum}`, text: text.trim() })
      }
      if (hymn.chorus?.trim()) selectedLyrics.push({ label: 'Chorus', text: hymn.chorus.trim() })

      return { name: sectionName, song: { title: hymn.name, lyrics: selectedLyrics } }
    })

  const formatDate = (isoDate) => {
    const d = new Date(isoDate + 'T00:00:00')
    return d.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
  }

  const handleDownload = async () => {
    setStatusMsg(null)
    setDownloading(true)
    try {
      const response = await fetch(API_URL, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          template_url: TEMPLATE_URL,
          date:         formatDate(date),
          sections:     buildSectionsPayload(),
        }),
      })
      if (!response.ok) {
        const err = await response.json().catch(() => ({ error: response.statusText }))
        throw new Error(err.error ?? `Server error ${response.status}`)
      }
      const blob = await response.blob()
      const url  = URL.createObjectURL(blob)
      const a    = document.createElement('a')
      a.href = url; a.download = 'Mass_Presentation.pptx'
      document.body.appendChild(a); a.click()
      document.body.removeChild(a); URL.revokeObjectURL(url)
      setStatusMsg({ type: 'success', text: 'Presentation generated successfully.' })
    } catch (err) {
      setStatusMsg({ type: 'error', text: `Generation failed: ${err.message}` })
    } finally {
      setDownloading(false)
    }
  }

  const canDownload = !downloading && !!date &&
    MASS_SECTIONS.some(s => selections[s].hymnId !== '')

  const selectedCount = MASS_SECTIONS.filter(s => selections[s].hymnId !== '').length

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  return (
    <div className="app">

      {/* ── Header ── */}
      <header className="header">
        <div className="header-eyebrow">Infant Jesus Church</div>
        <h1 className="header-title">Automated PPT<br />Generator</h1>
      </header>

      <main className="main">

        {/* ── Status Banner ── */}
        {(statusMsg || hymnsError) && (
          <div className={`banner banner--${hymnsError ? 'error' : statusMsg.type}`}>
            {(hymnsError || statusMsg?.type === 'error') ? <AlertIcon /> : <CheckIcon />}
            <span>{hymnsError ?? statusMsg?.text}</span>
          </div>
        )}

        {/* ── Date Card ── */}
        <section className="card" aria-label="Mass date">
          <div className="card-header">
            <CalendarIcon />
            <span className="card-label">Mass Date</span>
          </div>
          <div className="date-row">
            <label htmlFor="mass-date" className="date-label">
              Select the Sunday date for this Mass
            </label>
            <input
              id="mass-date"
              type="date"
              className="date-input"
              value={date}
              onChange={e => setDate(e.target.value)}
            />
          </div>
        </section>

        {/* ── Hymn Selections Card ── */}
        <section className="card" aria-label="Hymn selections">
          <div className="card-header">
            <MusicIcon />
            <span className="card-label">Hymn Selections</span>
            {selectedCount > 0 && (
              <span className="selection-badge">{selectedCount} of {MASS_SECTIONS.length}</span>
            )}
          </div>

          <div className="sections-list">
            {hymnsLoading
              ? Array.from({ length: 9 }).map((_, i) => (
                  <div key={i} className="skeleton-row">
                    <div className="skeleton skeleton--index" />
                    <div className="skeleton skeleton--select" />
                  </div>
                ))
              : visibleSections.map((name, index) => (
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
        </section>

        {/* ── Generate Button ── */}
        <div className="action-area">
          <button
            id="generate-ppt-btn"
            className="generate-btn"
            disabled={!canDownload}
            onClick={handleDownload}
          >
            {downloading ? (
              <>
                <span className="btn-spinner" />
                <span>Generating…</span>
              </>
            ) : (
              <>
                <DownloadIcon />
                <span>Generate Presentation</span>
              </>
            )}
          </button>
          {!canDownload && !downloading && (
            <p className="action-hint">Select a date and at least one hymn to continue</p>
          )}
        </div>

      </main>

      <footer className="footer">
        Clive Dsilva · Powered by Supabase &amp; python-pptx
      </footer>
    </div>
  )
}
