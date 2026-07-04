import { useState, useEffect, useCallback, useRef } from 'react'
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

// These sections auto-include ALL verses with no UI shown
const AUTO_VERSE_SECTIONS = new Set(['Lord Have Mercy', 'Gloria', 'Holy Holy', 'Proclamation'])

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

const API_URL = '/api/generate-ppt'

// ---------------------------------------------------------------------------
// Inline SVG icons
// ---------------------------------------------------------------------------

const ChevronIcon = ({ size = 14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
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

const SearchIcon = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="11" cy="11" r="8" />
    <line x1="21" y1="21" x2="16.65" y2="16.65" />
  </svg>
)

// ---------------------------------------------------------------------------
// HymnCombobox — searchable dropdown replacing native <select>
// ---------------------------------------------------------------------------

function HymnCombobox({ id, hymns, value, onChange }) {
  const [open,       setOpen]       = useState(false)
  const [query,      setQuery]      = useState('')
  const inputRef   = useRef(null)
  const wrapperRef = useRef(null)

  const selected = hymns.find(h => String(h.id) === String(value)) ?? null

  const filtered = query.trim()
    ? hymns.filter(h => h.name.toLowerCase().includes(query.toLowerCase()))
    : hymns

  useEffect(() => {
    const handler = (e) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target)) {
        setOpen(false)
        setQuery('')
      }
    }
    document.addEventListener('mousedown', handler)
    document.addEventListener('touchstart', handler)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('touchstart', handler)
    }
  }, [open])

  const openDropdown = () => {
    const rect = triggerRef.current?.getBoundingClientRect()
    if (!rect) return

    const PANEL_H  = 270
    const GAP      = 4
    const spaceBelow = window.innerHeight - rect.bottom
    const openDown   = spaceBelow >= PANEL_H || spaceBelow >= rect.top

    setPanelStyle(
      openDown
        ? { top: rect.bottom + GAP, left: rect.left, width: rect.width }
        : { bottom: window.innerHeight - rect.top + GAP, left: rect.left, width: rect.width }
    )
    setOpen(true)
    setQuery('')
    requestAnimationFrame(() => inputRef.current?.focus())
  }

  const selectHymn = (hymn) => {
    onChange(hymn ? String(hymn.id) : '')
    setOpen(false)
    setQuery('')
  }

  // Portal: render the panel directly on <body> so it's never clipped
  const panel = open && (
    <div
      id="combobox-portal"
      className="combobox-panel"
      role="listbox"
      style={{ position: 'fixed', zIndex: 9999, ...panelStyle }}
    >
      {/* Search bar at top */}
      <div className="combobox-search-row">
        <span className="combobox-search-icon"><SearchIcon /></span>
        <input
          ref={inputRef}
          className="combobox-search-input"
          type="text"
          placeholder="Search hymns…"
          value={query}
          onChange={e => setQuery(e.target.value)}
          autoComplete="off"
          autoCorrect="off"
          autoCapitalize="off"
          spellCheck={false}
        />
      </div>

      {/* Hymn list */}
      <div className="combobox-list">
        {/* Clear / None */}
        <div
          className={`combobox-option combobox-option--clear ${!value ? 'combobox-option--active' : ''}`}
          role="option"
          aria-selected={!value}
          onMouseDown={() => selectHymn(null)}
          onTouchEnd={(e) => { e.preventDefault(); selectHymn(null) }}
        >
          — None —
        </div>

        {filtered.length === 0
          ? <div className="combobox-empty">No hymns found</div>
          : filtered.map(h => (
              <div
                key={h.id}
                className={`combobox-option ${String(h.id) === String(value) ? 'combobox-option--active' : ''}`}
                role="option"
                aria-selected={String(h.id) === String(value)}
                onMouseDown={() => selectHymn(h)}
                onTouchEnd={(e) => { e.preventDefault(); selectHymn(h) }}
              >
                {h.name}
              </div>
            ))
        }
      </div>
    </div>
  )

  return (
    <>
      {/* Trigger — styled exactly like the original styled-select */}
      <div className="select-wrapper" ref={triggerRef}>
        <button
          type="button"
          id={id}
          className={`styled-select combobox-trigger-btn ${open ? 'combobox-trigger-btn--open' : ''}`}
          onClick={openDropdown}
          aria-haspopup="listbox"
          aria-expanded={open}
          style={{ textAlign: 'left', cursor: 'pointer' }}
        >
          <span className={selected ? '' : 'combobox-placeholder'}>
            {selected ? selected.name : '— Select hymn —'}
          </span>
        </button>
        <span className="select-chevron" style={{ pointerEvents: 'none' }}>
          <ChevronIcon />
        </span>
      </div>

      {/* Portal panel — rendered outside card, never clipped */}
      {panel}
    </>
  )
}

// ---------------------------------------------------------------------------
// VerseCountDropdown — compact "Up to verse N" selector
// ---------------------------------------------------------------------------

function VerseCountDropdown({ verseCount, upToVerse, onChange, hasChorus }) {
  const [open, setOpen] = useState(false)
  const wrapperRef = useRef(null)

  useEffect(() => {
    const handler = (e) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target)) setOpen(false)
    }
    document.addEventListener('mousedown', handler)
    document.addEventListener('touchstart', handler)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('touchstart', handler)
    }
  }, [])

  if (verseCount === 0) return null

  return (
    <div className="verse-count-row">
      <span className="verse-count-label">Verses</span>
      <div className="verse-count-picker" ref={wrapperRef}>
        <button
          type="button"
          className={`verse-count-trigger ${open ? 'verse-count-trigger--open' : ''}`}
          onClick={() => setOpen(o => !o)}
        >
          <span>{upToVerse}</span>
          <span className={`combobox-chevron ${open ? 'combobox-chevron--up' : ''}`}><ChevronIcon size={12} /></span>
        </button>
        {open && (
          <div className="verse-count-panel">
            {Array.from({ length: verseCount }, (_, i) => i + 1).map(n => (
              <div
                key={n}
                className={`verse-count-option ${n === upToVerse ? 'verse-count-option--active' : ''}`}
                onMouseDown={() => { onChange(n); setOpen(false) }}
                onTouchEnd={(e) => { e.preventDefault(); onChange(n); setOpen(false) }}
              >
                {n}
              </div>
            ))}
          </div>
        )}
      </div>
      {hasChorus && <span className="chorus-pill">♪ chorus</span>}
    </div>
  )
}

// ---------------------------------------------------------------------------
// SectionSelector — one row per Mass part
// ---------------------------------------------------------------------------

function SectionSelector({ index, name, hymns, hymnId, upToVerse, onChange }) {
  const filteredHymns = filterHymnsForSection(hymns, name)
  const selectedHymn  = hymns.find(h => String(h.id) === String(hymnId)) ?? null
  const verseCount    = selectedHymn?.verse_count ?? 0
  const hasChorus     = !!(selectedHymn?.chorus?.trim())
  const isAutoVerse   = AUTO_VERSE_SECTIONS.has(name)
  const isOptional    = name.endsWith('2') || name.endsWith('3') || name.endsWith('4')

  const handleHymnChange = (newId) => {
    if (!newId) { onChange(name, '', 0); return }
    const hymn = hymns.find(h => String(h.id) === String(newId))
    onChange(name, newId, hymn?.verse_count ?? 0)
  }

  return (
    <div className={`section-row ${hymnId ? 'section-row--active' : ''}`}>
      <div className="section-meta">
        <span className="section-index">{String(index + 1).padStart(2, '0')}</span>
        <div style={{ display: 'flex', flexDirection: 'column', marginTop: '-2px' }}>
          <span className="section-name">{name}</span>
          {isOptional && <span style={{ fontSize: '10px', color: 'var(--text-muted)', lineHeight: '1.2' }}>(Optional)</span>}
        </div>
      </div>

      <div className="section-right">
        <HymnCombobox
          id={`section-${index}`}
          hymns={filteredHymns}
          value={hymnId}
          onChange={handleHymnChange}
        />

        {hymnId && isAutoVerse && hasChorus && (
          <div style={{ marginTop: '4px' }}>
            <span className="chorus-pill">♪ chorus</span>
          </div>
        )}

        {hymnId && !isAutoVerse && (
          <VerseCountDropdown
            verseCount={verseCount}
            upToVerse={upToVerse}
            onChange={(n) => onChange(name, hymnId, n)}
            hasChorus={hasChorus}
          />
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
    Object.fromEntries(MASS_SECTIONS.map(s => [s, { hymnId: '', upToVerse: 0 }]))
  )

  const [downloading, setDownloading] = useState(false)
  const [statusMsg,   setStatusMsg]   = useState(null)

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

  const handleSelectionChange = useCallback((sectionName, hymnId, upToVerse) => {
    setSelections(prev => ({ ...prev, [sectionName]: { hymnId, upToVerse } }))
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
      const { hymnId, upToVerse } = selections[sectionName]
      const hymn = hymns.find(h => String(h.id) === String(hymnId))
      if (!hymn) return { name: sectionName, song: { title: '', lyrics: [] } }

      const isAuto = AUTO_VERSE_SECTIONS.has(sectionName)
      const versesToInclude = isAuto
        ? Array.from({ length: hymn.verse_count ?? 0 }, (_, i) => i + 1)
        : Array.from({ length: upToVerse }, (_, i) => i + 1)

      const selectedLyrics = []
      for (const verseNum of versesToInclude) {
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

  return (
    <div className="app">
      <header className="header">
        <div className="header-eyebrow">Infant Jesus Church</div>
        <h1 className="header-title">Automated PPT<br />Generator</h1>
      </header>

      <main className="main">
        {(statusMsg || hymnsError) && (
          <div className={`banner banner--${hymnsError ? 'error' : statusMsg.type}`}>
            {(hymnsError || statusMsg?.type === 'error') ? <AlertIcon /> : <CheckIcon />}
            <span>{hymnsError ?? statusMsg?.text}</span>
          </div>
        )}

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
                    upToVerse={selections[name].upToVerse}
                    onChange={handleSelectionChange}
                  />
                ))
            }
          </div>
        </section>

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
