import { useEffect, useState } from 'react'
import {
  AlertCircle,
  ArrowUpRight,
  Check,
  CircleHelp,
  Download,
  FileSpreadsheet,
  Link2,
  LoaderCircle,
  Trash2,
} from 'lucide-react'
import './App.css'

const MAX_LINKS = 50

function getDraftInfo(link, index) {
  try {
    const url = new URL(link)
    const parts = url.pathname.split('/').filter(Boolean)
    const draftId = parts.at(-1) || `Draft ${index + 1}`
    const game = url.searchParams.get('game')
    return {
      draftId,
      game: game ? `Game ${game}` : 'Draft',
      valid: ['drafter.lol', 'www.drafter.lol'].includes(url.hostname) && url.pathname.startsWith('/draft/'),
    }
  } catch {
    return { draftId: `Lien ${index + 1}`, game: 'URL invalide', valid: false }
  }
}

function App() {
  const [source, setSource] = useState('')
  const [apiState, setApiState] = useState('checking')
  const [isExporting, setIsExporting] = useState(false)
  const [message, setMessage] = useState(null)

  const links = source.split(/\r?\n/).map((link) => link.trim()).filter(Boolean)
  const drafts = links.map(getDraftInfo)
  const invalidCount = drafts.filter((draft) => !draft.valid).length

  useEffect(() => {
    const controller = new AbortController()
    fetch('/api/health', { signal: controller.signal })
      .then((response) => {
        if (!response.ok) throw new Error('API indisponible')
        setApiState('online')
      })
      .catch(() => {
        if (!controller.signal.aborted) setApiState('offline')
      })

    return () => controller.abort()
  }, [])

  async function handleExport(event) {
    event.preventDefault()
    if (!links.length || invalidCount || links.length > MAX_LINKS || isExporting) return

    setIsExporting(true)
    setMessage(null)

    try {
      const response = await fetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ urls: links }),
      })

      if (!response.ok) {
        const result = await response.json().catch(() => null)
        const detail = Array.isArray(result?.detail)
          ? result.detail.map((item) => item.msg).join(', ')
          : result?.detail
        throw new Error(detail || 'La génération du classeur a échoué.')
      }

      const blob = await response.blob()
      const filename = response.headers.get('Content-Disposition')?.match(/filename="?([^";]+)"?/i)?.[1]
        || 'drafter_drafts.xlsx'
      const downloadUrl = URL.createObjectURL(blob)
      const anchor = document.createElement('a')
      anchor.href = downloadUrl
      anchor.download = filename
      document.body.append(anchor)
      anchor.click()
      anchor.remove()
      window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000)
      setMessage({ type: 'success', text: `${links.length} draft${links.length > 1 ? 's' : ''} exporté${links.length > 1 ? 's' : ''}.` })
    } catch (error) {
      setMessage({ type: 'error', text: error.message || 'Impossible de contacter l’API.' })
    } finally {
      setIsExporting(false)
    }
  }

  const buttonDisabled = !links.length || invalidCount > 0 || links.length > MAX_LINKS || apiState !== 'online' || isExporting

  return (
    <div className="app-shell">
      <header className="topbar">
        <a className="wordmark" href="#top" aria-label="Draftsheet, accueil">
          <span className="wordmark-mark">D<span>/</span></span>
          <span className="wordmark-copy">DRAFTSHEET<small>LEAGUE TOOLING</small></span>
        </a>
        <div className={`api-indicator api-${apiState}`} role="status">
          <span className="api-dot" />
          {apiState === 'online' ? 'API en ligne' : apiState === 'checking' ? 'Connexion…' : 'API hors ligne'}
        </div>
      </header>

      <main id="top">
        <section className="intro">
          <div className="intro-kicker"><span /> DRAFTER.LOL <span className="kicker-rule" /> EXPORT XLSX</div>
          <h1>Rassemblez vos drafts<span>.</span></h1>
          <p>Importez jusqu’à 50 parties. Tous les drafts seront regroupés dans une seule feuille Excel.</p>
        </section>

        <form className="export-layout" onSubmit={handleExport}>
          <section className="input-panel" aria-labelledby="links-title">
            <div className="panel-heading">
              <div>
                <div className="eyebrow">SOURCE</div>
                <h2 id="links-title">Liens des drafts</h2>
              </div>
              <span className="count-pill"><Link2 size={15} /> {links.length} / {MAX_LINKS}</span>
            </div>

            <label className="field-label" htmlFor="draft-links">URL Drafter.lol</label>
            <textarea
              id="draft-links"
              value={source}
              onChange={(event) => {
                setSource(event.target.value)
                setMessage(null)
              }}
              placeholder={'https://drafter.lol/draft/PfemXp3j?game=1\nhttps://drafter.lol/draft/…?game=2'}
              spellCheck="false"
              autoCapitalize="off"
              autoCorrect="off"
              aria-describedby="input-hint"
            />
            <div className="field-meta" id="input-hint">
              <span>Un lien par ligne</span>
              <span>Maximum {MAX_LINKS}</span>
            </div>

            {invalidCount > 0 && (
              <p className="inline-warning"><AlertCircle size={15} /> Vérifiez les liens Drafter.lol invalides.</p>
            )}
            {links.length > MAX_LINKS && (
              <p className="inline-warning"><AlertCircle size={15} /> Réduisez la sélection à {MAX_LINKS} liens.</p>
            )}

            <div className="form-actions">
              <button
                className="clear-button"
                type="button"
                title="Effacer les liens"
                aria-label="Effacer les liens"
                disabled={!source || isExporting}
                onClick={() => {
                  setSource('')
                  setMessage(null)
                }}
              >
                <Trash2 size={17} />
              </button>
              <button className="export-button" type="submit" disabled={buttonDisabled}>
                {isExporting ? <LoaderCircle className="spin" size={18} /> : <Download size={18} />}
                {isExporting ? 'Génération…' : 'Télécharger Excel'}
                {!isExporting && <ArrowUpRight size={16} />}
              </button>
            </div>

            {message && (
              <div className={`result-message result-${message.type}`} role={message.type === 'error' ? 'alert' : 'status'}>
                {message.type === 'success' ? <Check size={17} /> : <AlertCircle size={17} />}
                <span>{message.text}</span>
              </div>
            )}
          </section>

          <aside className="queue-panel" aria-labelledby="queue-title">
            <div className="queue-heading">
              <div>
                <div className="eyebrow">À TRAITER</div>
                <h2 id="queue-title">File d’export</h2>
              </div>
              <FileSpreadsheet className="queue-icon" size={21} />
            </div>

            {drafts.length ? (
              <ol className="draft-list">
                {drafts.map((draft, index) => (
                  <li className={`draft-row ${draft.valid ? '' : 'draft-invalid'}`} key={`${links[index]}-${index}`}>
                    <span className="draft-index">{String(index + 1).padStart(2, '0')}</span>
                    <span className="draft-details">
                      <strong>{draft.draftId}</strong>
                      <small>{draft.game}</small>
                    </span>
                    <span className="draft-state">{draft.valid ? 'PRÊT' : 'À VÉRIFIER'}</span>
                  </li>
                ))}
              </ol>
            ) : (
              <div className="empty-state">
                <div className="empty-icon"><Link2 size={21} /></div>
                <p>La file est vide</p>
                <span>Les drafts ajoutés apparaîtront ici.</span>
              </div>
            )}

            <div className="queue-footnote">
              <CircleHelp size={15} />
              <span>Le classeur garde l’ordre des liens saisis.</span>
            </div>
          </aside>
        </form>
      </main>

      <footer className="footer-bar">
        <span>DRAFTSHEET <b>·</b> EXPORT LOCAL</span>
        <span>FORMAT <b>.XLSX</b></span>
      </footer>
    </div>
  )
}

export default App
