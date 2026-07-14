import { useState, useEffect, useCallback, useRef, useMemo } from 'react'
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

// Used in legacy mode only
const filterHymnsForSection = (hymns, category) =>
  hymns.filter(h => (h.categories ?? '').toLowerCase() === category.toLowerCase())

const TEMPLATE_URL =
  import.meta.env.VITE_TEMPLATE_URL ||
  'https://wvmxlnwfjtesbppojstu.supabase.co/storage/v1/object/public/templates/standard_template.pptx'

const API_URL = '/api/generate-ppt'

const CATEGORY_OPTIONS = [
  'entrance', 'lord have mercy', 'gloria', 'acclamation',
  'offertory', 'holy holy', 'proclamation', 'communion', 'recessional',
]

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

const PlusIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <line x1="12" y1="5" x2="12" y2="19" />
    <line x1="5" y1="12" x2="19" y2="12" />
  </svg>
)

const CloseIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <line x1="18" y1="6" x2="6" y2="18" />
    <line x1="6" y1="6" x2="18" y2="18" />
  </svg>
)

// ---------------------------------------------------------------------------
// HymnCombobox — searchable dropdown replacing native <select>
// ---------------------------------------------------------------------------

function HymnCombobox({ id, hymns, value, onChange }) {
  const [open,       setOpen]       = useState(false)
  const [query,      setQuery]      = useState('')
  const [panelStyle, setPanelStyle] = useState({})
  const inputRef   = useRef(null)
  const triggerRef = useRef(null)
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
  }

  const selectHymn = (hymn) => {
    onChange(hymn ? String(hymn.id) : '')
    setOpen(false)
    setQuery('')
  }

  // touchStartY tracks scroll vs tap — only select if finger moved < 8px
  const touchStartY = useRef(0)

  const panel = open && (
    <div
      id="combobox-portal"
      className="combobox-panel"
      role="listbox"
      style={{ position: 'fixed', zIndex: 9999, ...panelStyle }}
    >
      {/* Search bar */}
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
          onTouchStart={(e) => { touchStartY.current = e.touches[0].clientY }}
          onTouchEnd={(e) => {
            if (Math.abs(e.changedTouches[0].clientY - touchStartY.current) < 8) {
              e.preventDefault()
              selectHymn(null)
            }
          }}
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
                onTouchStart={(e) => { touchStartY.current = e.touches[0].clientY }}
                onTouchEnd={(e) => {
                  if (Math.abs(e.changedTouches[0].clientY - touchStartY.current) < 8) {
                    e.preventDefault()
                    selectHymn(h)
                  }
                }}
              >
                {h.name}
              </div>
            ))
        }
      </div>
    </div>
  )

  return (
    <div ref={wrapperRef}>
      {/* Trigger — styled like original styled-select */}
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
    </div>
  )
}

// ---------------------------------------------------------------------------
// VerseCountDropdown — compact "Up to verse N" selector
// ---------------------------------------------------------------------------

function VerseCountDropdown({ verseCount, upToVerse, onChange }) {
  const [open,       setOpen]      = useState(false)
  const [panelStyle, setPanelStyle] = useState({})
  const triggerRef = useRef(null)
  const touchStartY = useRef(0)

  useEffect(() => {
    if (!open) return
    const handler = (e) => {
      if (!triggerRef.current?.contains(e.target) &&
          !document.getElementById('verse-count-portal')?.contains(e.target)) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', handler)
    document.addEventListener('touchstart', handler)
    return () => {
      document.removeEventListener('mousedown', handler)
      document.removeEventListener('touchstart', handler)
    }
  }, [open])

  if (verseCount === 0) return null

  const openPicker = () => {
    const rect = triggerRef.current?.getBoundingClientRect()
    if (!rect) return
    const PANEL_H  = verseCount * 36 + 8
    const GAP      = 4
    const spaceBelow = window.innerHeight - rect.bottom
    const openDown   = spaceBelow >= PANEL_H || spaceBelow >= rect.top
    setPanelStyle(
      openDown
        ? { top: rect.bottom + GAP, left: rect.left, width: rect.width }
        : { bottom: window.innerHeight - rect.top + GAP, left: rect.left, width: rect.width }
    )
    setOpen(o => !o)
  }

  return (
    <div className="verse-count-row">
      <span className="verse-count-label">Verses</span>
      <div className="verse-count-picker" ref={triggerRef}>
        <button
          type="button"
          className={`verse-count-trigger ${open ? 'verse-count-trigger--open' : ''}`}
          onClick={openPicker}
        >
          <span>{upToVerse}</span>
          <span className={`combobox-chevron ${open ? 'combobox-chevron--up' : ''}`}><ChevronIcon size={12} /></span>
        </button>
      </div>

      {open && (
        <div
          id="verse-count-portal"
          className="verse-count-panel"
          style={{ position: 'fixed', zIndex: 9999, ...panelStyle }}
        >
          {Array.from({ length: verseCount }, (_, i) => i + 1).map(n => (
            <div
              key={n}
              className={`verse-count-option ${n === upToVerse ? 'verse-count-option--active' : ''}`}
              onMouseDown={() => { onChange(n); setOpen(false) }}
              onTouchStart={(e) => { touchStartY.current = e.touches[0].clientY }}
              onTouchEnd={(e) => {
                if (Math.abs(e.changedTouches[0].clientY - touchStartY.current) < 8) {
                  e.preventDefault()
                  onChange(n)
                  setOpen(false)
                }
              }}
            >
              {n}
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------
// AddHymnModal
// ---------------------------------------------------------------------------

function AddHymnModal({ onClose, onAdded }) {
  const [name,         setName]         = useState('')
  const [category,     setCategory]     = useState('')
  const [chorus,       setChorus]       = useState('')
  const [chorusBefore, setChorusBefore] = useState(false) // No = false, Yes = true
  const [verses,       setVerses]       = useState(['', '', ''])
  const [saving,       setSaving]       = useState(false)
  const [error,        setError]        = useState(null)
  const [success,      setSuccess]      = useState(false)

  // Lock body scroll while open
  useEffect(() => {
    document.body.style.overflow = 'hidden'
    return () => { document.body.style.overflow = '' }
  }, [])

  // Close on Escape
  useEffect(() => {
    const handler = (e) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])

  const updateVerse = (idx, val) =>
    setVerses(prev => prev.map((v, i) => i === idx ? val : v))

  const addVerseField    = () => setVerses(prev => [...prev, ''])
  const removeVerseField = (idx) => setVerses(prev => prev.filter((_, i) => i !== idx))

  const handleSave = async () => {
    setError(null)
    if (!name.trim()) { setError('Hymn name is required.'); return }
    if (!category)    { setError('Please select a category.'); return }

    const filledVerses = verses.map(v => v.trim())
    const verseCount   = filledVerses.filter(Boolean).length
    const chorusText   = chorus.trim()

    const payload = {
      name:        name.trim(),
      categories:  category,
      verse_count: verseCount,
    }

    if (!chorusBefore) {
      // Standard: chorus stored as-is, verses stored as-is
      payload.chorus = chorusText
      filledVerses.forEach((v, i) => {
        if (i < 5) payload[`verse_${i + 1}`] = v || null
      })
    } else {
      // Chorus-first: chorus column empty; prepend **chorus**\n\n to each verse
      payload.chorus = ''
      filledVerses.forEach((v, i) => {
        if (i < 5) {
          payload[`verse_${i + 1}`] = v ? `**${chorusText}**\n\n${v}` : null
        }
      })
    }

    setSaving(true)
    try {
      const { error: supaErr } = await supabase.from('hymns').insert([payload])
      if (supaErr) throw new Error(supaErr.message)
      setSuccess(true)
      setTimeout(() => { onAdded(); onClose() }, 1200)
    } catch (err) {
      setError(err.message)
    } finally {
      setSaving(false)
    }
  }

  return (
    <div className="modal-backdrop" onClick={(e) => { if (e.target === e.currentTarget) onClose() }}>
      <div className="modal" role="dialog" aria-modal="true" aria-label="Add new hymn">

        {/* Header */}
        <div className="modal-header">
          <span className="modal-title">Add New Hymn</span>
          <button type="button" className="modal-close" onClick={onClose} aria-label="Close">
            <CloseIcon />
          </button>
        </div>

        {/* Body */}
        <div className="modal-body">
          {error   && <div className="modal-banner modal-banner--error"><AlertIcon /><span>{error}</span></div>}
          {success && <div className="modal-banner modal-banner--success"><CheckIcon /><span>Hymn saved successfully!</span></div>}

          {/* Name */}
          <div className="modal-field">
            <label className="modal-label">HYMN NAME</label>
            <input
              className="modal-input"
              type="text"
              placeholder="e.g. Here I Am Lord"
              value={name}
              onChange={e => setName(e.target.value)}
            />
          </div>

          {/* Category */}
          <div className="modal-field">
            <label className="modal-label">CATEGORY</label>
            <div className="select-wrapper">
              <select
                className="styled-select"
                value={category}
                onChange={e => setCategory(e.target.value)}
              >
                <option value="">— Select category —</option>
                {CATEGORY_OPTIONS.map(c => (
                  <option key={c} value={c}>
                    {c.charAt(0).toUpperCase() + c.slice(1)}
                  </option>
                ))}
              </select>
              <span className="select-chevron"><ChevronIcon /></span>
            </div>
          </div>

          {/* Verses */}
          <div className="modal-field">
            <label className="modal-label">VERSES</label>
            {verses.map((v, i) => (
              <div key={i} className="modal-verse-row">
                <div className="modal-verse-header">
                  <span className="modal-verse-num">Verse {i + 1}</span>
                  {verses.length > 1 && (
                    <button
                      type="button"
                      className="modal-verse-remove"
                      onClick={() => removeVerseField(i)}
                      aria-label={`Remove verse ${i + 1}`}
                    >✕</button>
                  )}
                </div>
                <textarea
                  className="modal-textarea"
                  rows={4}
                  placeholder={`Type verse ${i + 1} lyrics here…`}
                  value={v}
                  onChange={e => updateVerse(i, e.target.value)}
                />
              </div>
            ))}
            {verses.length < 5 && (
              <button type="button" className="modal-add-verse" onClick={addVerseField}>
                <PlusIcon /> Add Verse
              </button>
            )}
          </div>

          {/* IS THE CHORUS BEFORE A VERSE? */}
          <div className="modal-field">
            <label className="modal-label">IS THE CHORUS BEFORE A VERSE?</label>
            <p className="modal-helper">
              (Click Yes if it is sung before the verses, click No if sung after)
            </p>
            <div className="modal-toggle-row">
              <button
                type="button"
                className={`modal-toggle-btn ${!chorusBefore ? 'modal-toggle-btn--active' : ''}`}
                onClick={() => setChorusBefore(false)}
              >No</button>
              <button
                type="button"
                className={`modal-toggle-btn ${chorusBefore ? 'modal-toggle-btn--active' : ''}`}
                onClick={() => setChorusBefore(true)}
              >Yes</button>
            </div>
          </div>

          {/* Chorus */}
          <div className="modal-field">
            <label className="modal-label">
              CHORUS <span className="modal-label-optional">(optional)</span>
            </label>
            <textarea
              className="modal-textarea"
              rows={4}
              placeholder="Type chorus lyrics here…"
              value={chorus}
              onChange={e => setChorus(e.target.value)}
            />
          </div>
        </div>

        {/* Footer */}
        <div className="modal-footer">
          <button type="button" className="modal-cancel" onClick={onClose} disabled={saving}>
            Cancel
          </button>
          <button type="button" className="modal-save" onClick={handleSave} disabled={saving || success}>
            {saving ? <><span className="btn-spinner" /> Saving…</> : 'Save Hymn'}
          </button>
        </div>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------
// SectionSelector — one row per Mass part
// ---------------------------------------------------------------------------

function SectionSelector({ index, label, category, isOptional, hymns, hymnId, onChange }) {
  const filteredHymns = filterHymnsForSection(hymns, category)

  const handleHymnChange = (newId) => {
    if (!newId) { onChange(label, ''); return }
    onChange(label, newId)
  }

  return (
    <div className={`section-row ${hymnId ? 'section-row--active' : ''}`}>
      <div className="section-meta">
        <span className="section-index">{String(index + 1).padStart(2, '0')}</span>
        <div style={{ display: 'flex', flexDirection: 'column', marginTop: '-2px' }}>
          <span className="section-name">{label}</span>
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
  const [showAddHymn,  setShowAddHymn]  = useState(false)

  const [date, setDate] = useState(() => {
    const today = new Date()
    const diff  = today.getDay() === 0 ? 0 : 7 - today.getDay()
    const next  = new Date(today)
    next.setDate(today.getDate() + diff)
    return next.toISOString().split('T')[0]
  })

  const [selections, setSelections] = useState(() =>
    Object.fromEntries(MASS_SECTIONS.map(s => [s, { hymnId: '' }]))
  )

  const [downloading,       setDownloading]       = useState(false)
  const [statusMsg,         setStatusMsg]         = useState(null)
  const [templates,         setTemplates]         = useState([])
  const [selectedTemplate,  setSelectedTemplate]  = useState(null)

  // Reload hymns list (called after successful insert)
  const reloadHymns = useCallback(() => {
    setHymnsLoading(true)
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

  useEffect(() => { reloadHymns() }, [reloadHymns])

  // Fetch templates from Supabase (include new dynamic-layout columns)
  useEffect(() => {
    supabase
      .from('templates')
      .select('id, name, file_url, structure, formatting_mode, fixed_font_size')
      .order('id', { ascending: true })
      .then(({ data, error }) => {
        if (error) {
          console.error('[templates] Fetch error (new columns may not exist yet):', error.message)
          // Fallback: fetch without new columns so tiles still show
          supabase
            .from('templates')
            .select('id, name, file_url')
            .order('id', { ascending: true })
            .then(({ data: d2 }) => {
              if (d2 && d2.length > 0) {
                setTemplates(d2)
                const def = d2.find(t => t.name.toLowerCase().includes('regular sunday mass')) ?? d2[0]
                setSelectedTemplate(def)
              }
            })
          return
        }
        if (data && data.length > 0) {
          console.log('[templates] Loaded:', data.map(t => ({
            id: t.id, name: t.name,
            hasStructure: !!t.structure,
            mode: t.formatting_mode,
            font: t.fixed_font_size,
          })))
          setTemplates(data)
          // Default to Regular Sunday Mass; fall back to first template if not found
          const def = data.find(t => t.name.toLowerCase().includes('regular sunday mass')) ?? data[0]
          setSelectedTemplate(def)
        }
      })
  }, [])

  // Derive the list of hymn slots from the selected template.
  // Dynamic mode: read from template.structure (only 'hymn' items).
  // Legacy mode:  fall back to the hardcoded MASS_SECTIONS list.
  const activeSections = useMemo(() => {
    if (selectedTemplate?.structure) {
      return selectedTemplate.structure
        .filter(item => item.type === 'hymn')
        .map(item => ({
          label:      item.label,
          category:   item.category,
          isOptional: false,
        }))
    }
    return MASS_SECTIONS.map(name => ({
      label:      name,
      category:   SECTION_TO_CATEGORY[name] ?? name.toLowerCase(),
      isOptional: name.endsWith('2') || name.endsWith('3') || name.endsWith('4'),
    }))
  }, [selectedTemplate])

  // Reset hymn selections whenever the active template changes.
  useEffect(() => {
    setSelections(
      Object.fromEntries(activeSections.map(s => [s.label, { hymnId: '' }]))
    )
  }, [selectedTemplate]) // eslint-disable-line react-hooks/exhaustive-deps

  const handleSelectionChange = useCallback((label, hymnId) => {
    setSelections(prev => ({ ...prev, [label]: { hymnId } }))
  }, [])

  const isSectionVisible = (label) => {
    // In dynamic mode all hymn slots are always visible
    if (selectedTemplate?.structure) return true
    // Legacy mode: optional duplicate sections hide until their primary is filled
    if (label === 'Entrance 2')    return !!selections['Entrance 1']?.hymnId
    if (label === 'Offertory 2')   return !!selections['Offertory 1']?.hymnId
    if (label === 'Communion 4')   return !!selections['Communion 3']?.hymnId
    if (label === 'Recessional 2') return !!selections['Recessional 1']?.hymnId
    return true
  }

  const visibleSections = activeSections.filter(sec => isSectionVisible(sec.label))

  const buildSectionsPayload = () =>
    activeSections.map(sec => {
      const hymnId = selections[sec.label]?.hymnId ?? ''
      const hymn   = hymns.find(h => String(h.id) === String(hymnId))
      if (!hymn) return { name: sec.label, song: { title: '', lyrics: [] } }

      // Always include ALL available verses
      const versesToInclude = Array.from({ length: hymn.verse_count ?? 0 }, (_, i) => i + 1)

      const selectedLyrics = []
      for (const verseNum of versesToInclude) {
        const text = hymn[`verse_${verseNum}`]
        if (text?.trim()) selectedLyrics.push({ label: `Verse ${verseNum}`, text: text.trim() })
      }
      if (hymn.chorus?.trim()) selectedLyrics.push({ label: 'Chorus', text: hymn.chorus.trim() })

      return { name: sec.label, song: { title: hymn.name, lyrics: selectedLyrics } }
    })

  const formatDate = (isoDate) => {
    const d = new Date(isoDate + 'T00:00:00')
    return d.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
  }

  const handleDownload = async () => {
    setStatusMsg(null)
    setDownloading(true)
    // Debug: log exactly what we are sending to the backend
    console.log('[generate] selectedTemplate:', {
      id:             selectedTemplate?.id,
      name:           selectedTemplate?.name,
      hasStructure:   !!selectedTemplate?.structure,
      formattingMode: selectedTemplate?.formatting_mode,
      fontSize:       selectedTemplate?.fixed_font_size,
    })
    try {
      const response = await fetch(API_URL, {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          template_url:    selectedTemplate?.file_url ?? TEMPLATE_URL,
          date:            formatDate(date),
          structure:       selectedTemplate?.structure       ?? null,
          formatting_mode: selectedTemplate?.formatting_mode ?? 'auto_fit',
          fixed_font_size: selectedTemplate?.fixed_font_size ?? 36,
          sections:        buildSectionsPayload(),
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
    visibleSections.some(s => selections[s.label]?.hymnId !== '')

  const selectedCount = visibleSections.filter(s => selections[s.label]?.hymnId !== '').length

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

        {/* Template Selector */}
        {templates.length > 0 && (
          <section className="card template-card" aria-label="Template selector">
            <div className="card-header">
              <span className="card-label">Select Theme</span>
            </div>
            <div className="template-grid">
              {templates.map(t => (
                <button
                  key={t.id}
                  type="button"
                  className={`template-tile ${
                    selectedTemplate?.id === t.id ? 'template-tile--active' : ''
                  }`}
                  onClick={() => setSelectedTemplate(t)}
                >
                  {t.name}
                </button>
              ))}
            </div>
          </section>
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
              <span className="selection-badge">{selectedCount} of {visibleSections.length}</span>
            )}
            <button
              type="button"
              className="add-hymn-btn"
              onClick={() => setShowAddHymn(true)}
              title="Add a new hymn to the database"
            >
              <PlusIcon /> Add Hymn
            </button>
          </div>

          <div className="sections-list">
            {hymnsLoading
              ? Array.from({ length: 9 }).map((_, i) => (
                  <div key={i} className="skeleton-row">
                    <div className="skeleton skeleton--index" />
                    <div className="skeleton skeleton--select" />
                  </div>
                ))
              : visibleSections.map((sec, index) => (
                  <SectionSelector
                    key={sec.label}
                    index={index}
                    label={sec.label}
                    category={sec.category}
                    isOptional={sec.isOptional}
                    hymns={hymns}
                    hymnId={selections[sec.label]?.hymnId ?? ''}
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

      {showAddHymn && (
        <AddHymnModal
          onClose={() => setShowAddHymn(false)}
          onAdded={reloadHymns}
        />
      )}
    </div>
  )
}
