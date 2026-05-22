import { useState, useCallback, useRef } from 'react'
import * as XLSX from 'xlsx'
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  AreaChart, Area,
} from 'recharts'

// ── Paleta de colores ──────────────────────────────────────────────
const C = ['#1d4ed8','#059669','#d97706','#dc2626','#7c3aed','#0891b2','#65a30d','#db2777','#6b7280','#0f766e']

// ── Helpers ────────────────────────────────────────────────────────
const excelToDate = (v) => {
  if (v == null || v === '') return null
  if (v instanceof Date) return v
  const n = Number(v)
  if (!isNaN(n) && n > 40000 && n < 55000) {
    const d = new Date(Math.round((n - 25569) * 86400000))
    return isNaN(d.getTime()) ? null : d
  }
  if (typeof v === 'string' && v.length > 5) {
    const d = new Date(v.replace(/\//g, '-'))
    return isNaN(d.getTime()) ? null : d
  }
  return null
}

const fmtMonth = (d) => !d ? '' : d.toLocaleDateString('es', { month: 'short', year: '2-digit' })
const f1 = (n) => n == null ? '—' : parseFloat(Number(n).toFixed(1))
const f0 = (n) => n == null ? '—' : Math.round(Number(n))

// ── Definición de indicadores ──────────────────────────────────────
const INDICATORS = [
  { id: 'kpis',        label: 'KPIs generales',               group: 'Resumen',      icon: 'ti-dashboard' },
  { id: 'byTipo',      label: 'Por tipo de brigada',           group: 'Personal',     icon: 'ti-users-group' },
  { id: 'byArea',      label: 'Por área de trabajo',           group: 'Personal',     icon: 'ti-building' },
  { id: 'byGrupo',     label: 'Por grupo de trabajo',          group: 'Personal',     icon: 'ti-hierarchy' },
  { id: 'timeline',    label: 'Actividad mensual',             group: 'Temporal',     icon: 'ti-calendar-stats' },
  { id: 'modHoras',    label: 'Horas por módulo',              group: 'Programa',     icon: 'ti-clock-hour-3' },
  { id: 'modPers',     label: 'Participantes por módulo',      group: 'Programa',     icon: 'ti-users' },
  { id: 'tipoAct',     label: 'Tipo de actividad',             group: 'Programa',     icon: 'ti-list-check' },
  { id: 'notasMod',    label: 'Notas promedio por módulo',     group: 'Evaluaciones', icon: 'ti-star' },
  { id: 'distNotas',   label: 'Distribución de notas',         group: 'Evaluaciones', icon: 'ti-chart-bar' },
  { id: 'cumplim',     label: 'Cumplimiento del programa',     group: 'Cumplimiento', icon: 'ti-circle-check' },
  { id: 'topBV',       label: 'Mayor participación',           group: 'Rankings',     icon: 'ti-award' },
  { id: 'botBV',       label: 'Menor participación',           group: 'Rankings',     icon: 'ti-alert-triangle' },
  { id: 'instructores',label: 'Por instructor',                group: 'Instructores', icon: 'ti-school' },
  { id: 'empresas',    label: 'Empresa instructora',           group: 'Instructores', icon: 'ti-building-factory-2' },
]

// ── Parser del workbook ────────────────────────────────────────────
function parseData(wb) {
  const j = (name) => {
    const s = wb.Sheets[name]
    return s ? XLSX.utils.sheet_to_json(s, { defval: null, raw: true }) : []
  }
  return { personal: j('BD_personal'), actividades: j('BD_seg'), catalogo: j('_tbl2') }
}

// ── Motor de estadísticas ──────────────────────────────────────────
function computeStats(raw) {
  const { personal = [], actividades = [], catalogo = [] } = raw

  /* 1 · Directorio de personal */
  const pMap = {}
  personal.forEach(r => {
    const ci = String(r['CEDULA IDENTIDAD'] || '').trim()
    if (!ci || ci === 'null' || ci === '0') return
    pMap[ci] = {
      ci,
      nombre: String(r['NOMBRE'] || '').trim(),
      tipo: String(r['BV- GPR-LE'] || '').trim(),
      area: String(r['Area_2'] || r['Area'] || '').trim(),
      estado: String(r['Estado_AP'] || '').trim(),
      grupo: String(r['GRUPO DE TRABAJO'] || '').trim(),
    }
  })

  /* 2 · Actividades */
  const acts = []
  actividades.forEach(r => {
    const ci = String(r['CEDULA IDENTIDAD'] || '').trim()
    if (!ci || ci === 'null') return
    const notaRaw = r['nota']
    const nota = notaRaw !== null && notaRaw !== '' && !isNaN(Number(notaRaw)) && Number(notaRaw) > 0
      ? Number(notaRaw) : null
    const tiempo = parseFloat(r['TIEMPO (hr)'] || 0) || 0
    const mod = String(r['Modulo'] || '').trim()
    acts.push({
      ci, nombre: String(r['NOMBRE'] || '').trim(),
      servicio: String(r['Servicio'] || '').trim(),
      modulo: mod === '0' ? '' : mod,
      nota, tiempo,
      fecha: excelToDate(r['FECHA(yyyy-mm-dd)']),
      instructor: String(r['Instructor'] || '').trim(),
      empresa: String(r['Empresa_inst'] || '').trim(),
    })
  })

  /* 3 · KPIs */
  const totalHoras = acts.reduce((s, a) => s + a.tiempo, 0)
  const actsNota = acts.filter(a => a.nota !== null)
  const notaGlobal = actsNota.length
    ? actsNota.reduce((s, a) => s + a.nota, 0) / actsNota.length : null

  /* 4 · Personal agrupado */
  const mkTipoArea = (items, keyFn, activeFn) => {
    const m = {}
    items.forEach(p => {
      const k = (keyFn(p) || '').trim() || 'Sin dato'
      if (k === 'Sin dato') return
      if (!m[k]) m[k] = { count: 0, activos: 0 }
      m[k].count++
      if (activeFn(p)) m[k].activos++
    })
    return Object.entries(m)
      .map(([k, v]) => ({ label: k, ...v }))
      .sort((a, b) => b.count - a.count)
  }

  const byTipo = mkTipoArea(Object.values(pMap), p => p.tipo, p => p.estado === 'Activo')
  const byArea = mkTipoArea(Object.values(pMap), p => p.area, p => p.estado === 'Activo')
  const byGrupo = mkTipoArea(Object.values(pMap), p => p.grupo, p => p.estado === 'Activo').slice(0, 14)

  /* 5 · Módulos */
  const modM = {}
  acts.forEach(a => {
    if (!a.modulo) return
    if (!modM[a.modulo]) modM[a.modulo] = { ses: 0, horas: 0, cis: new Set(), notas: [] }
    modM[a.modulo].ses++
    modM[a.modulo].horas += a.tiempo
    modM[a.modulo].cis.add(a.ci)
    if (a.nota !== null) modM[a.modulo].notas.push(a.nota)
  })
  const moduloStats = Object.entries(modM)
    .map(([m, s]) => ({
      modulo: m.replace('MODULO ', 'M'),
      moduloFull: m,
      sesiones: s.ses,
      horas: parseFloat(s.horas.toFixed(1)),
      personas: s.cis.size,
      nota: s.notas.length
        ? parseFloat((s.notas.reduce((a, b) => a + b, 0) / s.notas.length).toFixed(1)) : null,
    }))
    .sort((a, b) => a.moduloFull.localeCompare(b.moduloFull))

  /* 6 · Timeline mensual */
  const mesM = {}
  acts.forEach(a => {
    if (!a.fecha) return
    const k = `${a.fecha.getFullYear()}-${String(a.fecha.getMonth() + 1).padStart(2, '0')}`
    if (!mesM[k]) mesM[k] = { ses: 0, horas: 0, cis: new Set(), label: fmtMonth(a.fecha) }
    mesM[k].ses++
    mesM[k].horas += a.tiempo
    mesM[k].cis.add(a.ci)
  })
  const timeline = Object.entries(mesM)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([, v]) => ({ label: v.label, sesiones: v.ses, horas: parseFloat(v.horas.toFixed(1)), personas: v.cis.size }))

  /* 7 · Por persona */
  const perM = {}
  acts.forEach(a => {
    if (!a.ci) return
    if (!perM[a.ci]) perM[a.ci] = { horas: 0, ses: 0, mods: new Set(), notas: [], nombre: a.nombre }
    perM[a.ci].horas += a.tiempo
    perM[a.ci].ses++
    if (a.modulo) perM[a.ci].mods.add(a.modulo)
    if (a.nota !== null) perM[a.ci].notas.push(a.nota)
  })
  const personaStats = Object.entries(perM)
    .map(([ci, s]) => ({
      ci, nombre: s.nombre || pMap[ci]?.nombre || ci,
      horas: parseFloat(s.horas.toFixed(1)),
      ses: s.ses, mods: s.mods,
      nota: s.notas.length
        ? parseFloat((s.notas.reduce((a, b) => a + b, 0) / s.notas.length).toFixed(1)) : null,
      tipo: pMap[ci]?.tipo || '', area: pMap[ci]?.area || '', estado: pMap[ci]?.estado || '',
    }))
    .sort((a, b) => b.horas - a.horas)

  const topBV = personaStats.filter(p => p.horas > 0).slice(0, 10)
  const botBV = [...personaStats].filter(p => p.horas > 0).sort((a, b) => a.horas - b.horas).slice(0, 10)

  /* 8 · Instructores y empresas */
  const instM = {}; const empM = {}
  acts.forEach(a => {
    if (a.instructor) {
      if (!instM[a.instructor]) instM[a.instructor] = { ses: 0, horas: 0 }
      instM[a.instructor].ses++; instM[a.instructor].horas += a.tiempo
    }
    if (a.empresa) {
      if (!empM[a.empresa]) empM[a.empresa] = { ses: 0, horas: 0 }
      empM[a.empresa].ses++; empM[a.empresa].horas += a.tiempo
    }
  })
  const instructores = Object.entries(instM)
    .map(([k, v]) => ({ label: k, sesiones: v.ses, horas: parseFloat(v.horas.toFixed(1)) }))
    .sort((a, b) => b.sesiones - a.sesiones).slice(0, 8)
  const empresas = Object.entries(empM)
    .map(([k, v]) => ({ label: k, sesiones: v.ses, horas: parseFloat(v.horas.toFixed(1)) }))
    .sort((a, b) => b.sesiones - a.sesiones)

  /* 9 · Distribución de notas */
  const nr = { '< 60': 0, '60–69': 0, '70–79': 0, '80–89': 0, '90–100': 0 }
  acts.forEach(a => {
    if (a.nota === null) return
    if (a.nota < 60) nr['< 60']++
    else if (a.nota < 70) nr['60–69']++
    else if (a.nota < 80) nr['70–79']++
    else if (a.nota < 90) nr['80–89']++
    else nr['90–100']++
  })
  const distNotas = Object.entries(nr).map(([rango, count]) => ({ rango, count }))

  /* 10 · Tipo de actividad */
  const srvM = {}
  acts.forEach(a => { const s = a.servicio || 'Sin tipo'; if (!srvM[s]) srvM[s] = 0; srvM[s]++ })
  const tipoAct = Object.entries(srvM).map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count)

  /* 11 · Cumplimiento vs. programa requerido */
  const reqByTipo = {}
  catalogo.forEach(r => {
    const m = String(r['MODULO'] || '').trim()
    if (!m || m === '0') return
    ;['BV-B', 'BV-M', 'GPR', 'LE', 'URE-M'].forEach(t => {
      if (Number(r[t] || 0) > 0) {
        if (!reqByTipo[t]) reqByTipo[t] = new Set()
        reqByTipo[t].add(m)
      }
    })
  })
  const cumplimiento = Object.values(pMap)
    .filter(p => p.estado === 'Activo')
    .map(p => {
      const req = reqByTipo[p.tipo] || new Set()
      const reqArr = [...req]
      const done = perM[p.ci]?.mods || new Set()
      const completados = reqArr.filter(m => done.has(m)).length
      const requeridos = reqArr.length
      return {
        ci: p.ci, nombre: p.nombre, tipo: p.tipo, area: p.area,
        horas: parseFloat((perM[p.ci]?.horas || 0).toFixed(1)),
        requeridos, completados,
        pct: requeridos > 0 ? Math.round((completados / requeridos) * 100) : 0,
      }
    })
    .filter(p => p.requeridos > 0)
    .sort((a, b) => b.pct - a.pct)

  return {
    kpis: {
      totalPersonal: Object.keys(pMap).length,
      activos: Object.values(pMap).filter(p => p.estado === 'Activo').length,
      bajas: Object.values(pMap).filter(p => p.estado === 'Baja').length,
      totalHoras: parseFloat(totalHoras.toFixed(1)),
      totalSesiones: acts.length,
      modUnicos: Object.keys(modM).length,
      notaGlobal: notaGlobal !== null ? parseFloat(notaGlobal.toFixed(1)) : null,
      persCapacitadas: Object.keys(perM).length,
    },
    byTipo, byArea, byGrupo,
    moduloStats, timeline,
    topBV, botBV,
    instructores, empresas,
    distNotas, tipoAct, cumplimiento,
  }
}

// ── Tooltip personalizado ──────────────────────────────────────────
const TT = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null
  return (
    <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 8, padding: '8px 12px', fontSize: 12, boxShadow: '0 4px 12px rgba(0,0,0,0.08)' }}>
      {label && <p style={{ color: '#6b7280', margin: '0 0 4px', fontWeight: 500 }}>{label}</p>}
      {payload.map((p, i) => (
        <p key={i} style={{ color: p.color || p.fill || '#111', margin: '2px 0' }}>
          {p.name}: <strong>{typeof p.value === 'number' ? f1(p.value) : p.value}</strong>
        </p>
      ))}
    </div>
  )
}

// ── Componentes de visualización ───────────────────────────────────

const Card = ({ title, icon, children, full }) => (
  <div className="card" style={{
    background: '#fff', border: '1px solid #e5e7eb', borderRadius: 12,
    padding: '1rem 1.25rem', gridColumn: full ? '1 / -1' : undefined,
    boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
  }}>
    {title && (
      <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 14 }}>
        {icon && <i className={`ti ${icon}`} aria-hidden="true" style={{ fontSize: 15, color: '#6b7280' }} />}
        <span style={{ fontSize: 13, fontWeight: 500, color: '#374151' }}>{title}</span>
      </div>
    )}
    {children}
  </div>
)

const Kpi = ({ label, value, sub, color = '#111827' }) => (
  <div style={{ background: '#f8f9fa', borderRadius: 8, padding: '0.875rem 1rem' }}>
    <p style={{ fontSize: 11, color: '#6b7280', margin: '0 0 4px' }}>{label}</p>
    <p style={{ fontSize: 28, fontWeight: 600, margin: 0, color, lineHeight: 1.1 }}>{value ?? '—'}</p>
    {sub && <p style={{ fontSize: 11, color: '#9ca3af', margin: '2px 0 0' }}>{sub}</p>}
  </div>
)

function KPIsPanel({ d }) {
  return (
    <Card full icon="ti-dashboard" title="Resumen ejecutivo">
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 10 }}>
        <Kpi label="Personal registrado" value={f0(d.totalPersonal)} />
        <Kpi label="Personal activo" value={f0(d.activos)} color="#1d4ed8" />
        <Kpi label="Bajas" value={f0(d.bajas)} color="#dc2626" />
        <Kpi label="Personas capacitadas" value={f0(d.persCapacitadas)} color="#059669" />
        <Kpi label="Horas totales dictadas" value={f1(d.totalHoras)} sub="horas" />
        <Kpi label="Sesiones registradas" value={f0(d.totalSesiones)} />
        <Kpi label="Módulos activos" value={f0(d.modUnicos)} />
        <Kpi label="Nota global promedio" value={d.notaGlobal !== null ? f1(d.notaGlobal) : '—'} sub="sobre 100" color="#d97706" />
      </div>
    </Card>
  )
}

function PieCard({ title, icon, data, nameKey = 'label', valKey = 'count', colors = C }) {
  const total = data.reduce((s, r) => s + (r[valKey] || 0), 0)
  return (
    <Card icon={icon} title={title}>
      <div style={{ height: 190 }}>
        <ResponsiveContainer>
          <PieChart>
            <Pie data={data} dataKey={valKey} nameKey={nameKey} cx="50%" cy="50%" innerRadius={50} outerRadius={80} paddingAngle={2}>
              {data.map((_, i) => <Cell key={i} fill={colors[i % colors.length]} />)}
            </Pie>
            <Tooltip content={<TT />} />
          </PieChart>
        </ResponsiveContainer>
      </div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px 12px', marginTop: 8 }}>
        {data.map((r, i) => (
          <span key={i} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11, color: '#6b7280' }}>
            <span style={{ width: 9, height: 9, borderRadius: 2, background: colors[i % colors.length], flexShrink: 0 }} />
            {r[nameKey]} ({total > 0 ? Math.round(r[valKey] / total * 100) : 0}%)
          </span>
        ))}
      </div>
    </Card>
  )
}

function HBarCard({ title, icon, data, xKey, bars, colors = C, full }) {
  const h = Math.max(200, data.length * 38 + 60)
  return (
    <Card icon={icon} title={title} full={full}>
      <div style={{ height: h }}>
        <ResponsiveContainer>
          <BarChart data={data} layout="vertical" margin={{ left: 0, right: 16, top: 4, bottom: 4 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f3f4f6" />
            <XAxis type="number" tick={{ fontSize: 10, fill: '#9ca3af' }} />
            <YAxis type="category" dataKey={xKey} width={100} tick={{ fontSize: 11, fill: '#6b7280' }} />
            <Tooltip content={<TT />} />
            {bars.length > 1 && <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />}
            {bars.map((b, i) => (
              <Bar key={b.key} dataKey={b.key} name={b.label || b.key} fill={colors[i % colors.length]} radius={[0, 3, 3, 0]} maxBarSize={18} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Card>
  )
}

function VBarCard({ title, icon, data, xKey, bars, colors = C, full }) {
  return (
    <Card icon={icon} title={title} full={full}>
      <div style={{ height: 260 }}>
        <ResponsiveContainer>
          <BarChart data={data} margin={{ left: 0, right: 8, top: 4, bottom: 48 }}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f3f4f6" />
            <XAxis dataKey={xKey} tick={{ fontSize: 10, fill: '#9ca3af' }} angle={-40} textAnchor="end" interval={0} />
            <YAxis tick={{ fontSize: 10, fill: '#9ca3af' }} />
            <Tooltip content={<TT />} />
            {bars.length > 1 && <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />}
            {bars.map((b, i) => (
              <Bar key={b.key} dataKey={b.key} name={b.label || b.key} fill={colors[i % colors.length]} radius={[3, 3, 0, 0]} maxBarSize={24} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Card>
  )
}

function TimelineCard({ data }) {
  return (
    <Card full icon="ti-calendar-stats" title="Actividad mensual — sesiones y participantes únicos">
      <div style={{ height: 240 }}>
        <ResponsiveContainer>
          <AreaChart data={data} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
            <defs>
              <linearGradient id="g1" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C[0]} stopOpacity={0.25} />
                <stop offset="95%" stopColor={C[0]} stopOpacity={0} />
              </linearGradient>
              <linearGradient id="g2" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C[1]} stopOpacity={0.25} />
                <stop offset="95%" stopColor={C[1]} stopOpacity={0} />
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f3f4f6" />
            <XAxis dataKey="label" tick={{ fontSize: 10, fill: '#9ca3af' }} />
            <YAxis tick={{ fontSize: 10, fill: '#9ca3af' }} />
            <Tooltip content={<TT />} />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
            <Area type="monotone" dataKey="sesiones" name="Sesiones" stroke={C[0]} fill="url(#g1)" strokeWidth={2} dot={false} />
            <Area type="monotone" dataKey="personas" name="Participantes únicos" stroke={C[1]} fill="url(#g2)" strokeWidth={2} dot={false} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </Card>
  )
}

function CumplimientoCard({ data }) {
  const bins = [
    { label: '0–24%', color: '#dc2626', count: 0 },
    { label: '25–49%', color: '#d97706', count: 0 },
    { label: '50–74%', color: '#2563eb', count: 0 },
    { label: '75–99%', color: '#7c3aed', count: 0 },
    { label: '100%', color: '#059669', count: 0 },
  ]
  data.forEach(p => {
    if (p.pct < 25) bins[0].count++
    else if (p.pct < 50) bins[1].count++
    else if (p.pct < 75) bins[2].count++
    else if (p.pct < 100) bins[3].count++
    else bins[4].count++
  })

  const top6 = data.slice(0, 6)
  const bot6 = [...data].sort((a, b) => a.pct - b.pct).slice(0, 6)

  return (
    <Card full icon="ti-circle-check" title={`Cumplimiento del programa — ${data.length} brigadistas activos con programa definido`}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: 8, marginBottom: 20 }}>
        {bins.map((b, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 10, background: '#f8f9fa', padding: '10px 14px', borderRadius: 8 }}>
            <span style={{ fontSize: 24, fontWeight: 600, color: b.color }}>{b.count}</span>
            <div>
              <p style={{ margin: 0, fontSize: 12, fontWeight: 500 }}>{b.label}</p>
              <p style={{ margin: 0, fontSize: 11, color: '#9ca3af' }}>brigadistas</p>
            </div>
          </div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
        {[{ label: 'Mayor cumplimiento', list: top6, color: '#059669' }, { label: 'Menor cumplimiento', list: bot6, color: '#dc2626' }].map(({ label, list, color }) => (
          <div key={label}>
            <p style={{ fontSize: 12, fontWeight: 500, color: '#6b7280', margin: '0 0 10px' }}>{label}</p>
            {list.map((p, i) => (
              <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <span style={{ fontSize: 11, color: '#9ca3af', width: 18, textAlign: 'right' }}>{i + 1}.</span>
                <div style={{ flex: 1 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3 }}>
                    <span style={{ fontSize: 11, maxWidth: 130, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.nombre}</span>
                    <span style={{ fontSize: 11, fontWeight: 600, color }}>{p.pct}%</span>
                  </div>
                  <div style={{ background: '#f3f4f6', borderRadius: 3, height: 5, overflow: 'hidden' }}>
                    <div style={{ width: `${Math.max(p.pct, 2)}%`, height: '100%', background: color, borderRadius: 3 }} />
                  </div>
                  <span style={{ fontSize: 10, color: '#9ca3af' }}>{p.tipo} · {p.area || '—'}</span>
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>
    </Card>
  )
}

function RankingCard({ title, icon, data, high }) {
  const max = data.length ? Math.max(...data.map(d => d.horas)) : 1
  const color = high ? '#059669' : '#dc2626'
  return (
    <Card icon={icon} title={title}>
      {data.map((p, i) => (
        <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 9 }}>
          <span style={{ fontSize: 11, color: '#9ca3af', width: 18, textAlign: 'right' }}>{i + 1}.</span>
          <div style={{ flex: 1 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 2 }}>
              <span style={{ fontSize: 12, maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.nombre}</span>
              <span style={{ fontSize: 11, fontWeight: 600, color }}>{f1(p.horas)} h</span>
            </div>
            <div style={{ background: '#f3f4f6', borderRadius: 3, height: 4, overflow: 'hidden' }}>
              <div style={{ width: `${(p.horas / max) * 100}%`, height: '100%', background: color, borderRadius: 3 }} />
            </div>
            <span style={{ fontSize: 10, color: '#9ca3af' }}>{p.tipo} · {p.area || '—'}</span>
          </div>
        </div>
      ))}
    </Card>
  )
}

function DistNotasCard({ data }) {
  const cols = ['#dc2626', '#d97706', '#2563eb', '#7c3aed', '#059669']
  return (
    <Card icon="ti-chart-bar" title="Distribución de notas registradas">
      <div style={{ height: 200 }}>
        <ResponsiveContainer>
          <BarChart data={data} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f3f4f6" />
            <XAxis dataKey="rango" tick={{ fontSize: 11, fill: '#9ca3af' }} />
            <YAxis tick={{ fontSize: 10, fill: '#9ca3af' }} />
            <Tooltip content={<TT />} />
            <Bar dataKey="count" name="Evaluaciones" radius={[3, 3, 0, 0]} maxBarSize={40}>
              {data.map((_, i) => <Cell key={i} fill={cols[i]} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Card>
  )
}

// ── Upload Zone ────────────────────────────────────────────────────
function UploadZone({ onFile, loading, error, fileInfo, dragging, setDragging }) {
  const ref = useRef()
  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false)
    const f = e.dataTransfer.files[0]
    if (f) onFile(f)
  }, [onFile, setDragging])

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => ref.current?.click()}
      style={{
        border: `2px dashed ${dragging ? '#1d4ed8' : '#d1d5db'}`,
        borderRadius: 10, padding: '1.25rem', textAlign: 'center', cursor: 'pointer',
        background: dragging ? '#eff6ff' : '#f9fafb',
        transition: 'all 0.15s',
      }}>
      <input ref={ref} type="file" accept=".xlsx" style={{ display: 'none' }}
        onChange={e => { if (e.target.files[0]) onFile(e.target.files[0]) }} />
      {loading ? (
        <p style={{ fontSize: 13, color: '#6b7280', margin: 0 }}>Procesando archivo…</p>
      ) : fileInfo ? (
        <>
          <i className="ti ti-circle-check" style={{ fontSize: 22, color: '#059669', display: 'block' }} aria-hidden="true" />
          <p style={{ fontSize: 12, fontWeight: 500, margin: '6px 0 2px' }}>{fileInfo.name}</p>
          <p style={{ fontSize: 11, color: '#9ca3af', margin: 0 }}>{fileInfo.size} · Haz clic para cambiar</p>
        </>
      ) : (
        <>
          <i className="ti ti-file-spreadsheet" style={{ fontSize: 26, color: '#9ca3af', display: 'block' }} aria-hidden="true" />
          <p style={{ fontSize: 13, margin: '8px 0 2px', fontWeight: 500 }}>DB_CAPACITACION.xlsx</p>
          <p style={{ fontSize: 11, color: '#9ca3af', margin: 0 }}>Arrastra o haz clic para subir</p>
        </>
      )}
      {error && <p style={{ fontSize: 11, color: '#dc2626', margin: '8px 0 0' }}>{error}</p>}
    </div>
  )
}

// ── App principal ──────────────────────────────────────────────────
export default function App() {
  const [stats, setStats] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)
  const [fileInfo, setFileInfo] = useState(null)
  const [dragging, setDragging] = useState(false)
  const [selected, setSelected] = useState(['kpis', 'byTipo', 'byArea', 'timeline', 'modHoras', 'cumplim', 'topBV'])
  const [title, setTitle] = useState('Reporte semanal de capacitación')

  const handleFile = useCallback(async (file) => {
    setLoading(true); setError(null)
    try {
      const buf = await file.arrayBuffer()
      const wb = XLSX.read(buf, { type: 'array', raw: true })
      const raw = parseData(wb)
      setStats(computeStats(raw))
      setFileInfo({ name: file.name, size: `${(file.size / 1024).toFixed(0)} KB` })
    } catch (e) {
      setError('Error al procesar: ' + e.message)
    } finally {
      setLoading(false)
    }
  }, [])

  const toggle = (id) => setSelected(p => p.includes(id) ? p.filter(i => i !== id) : [...p, id])
  const groups = [...new Set(INDICATORS.map(i => i.group))]

  const render = (id) => {
    if (!stats) return null
    const { kpis, byTipo, byArea, byGrupo, moduloStats, timeline,
      topBV, botBV, instructores, empresas, distNotas, tipoAct, cumplimiento } = stats

    switch (id) {
      case 'kpis': return <KPIsPanel key={id} d={kpis} />
      case 'byTipo': return <PieCard key={id} title="Por tipo de brigada" icon="ti-users-group" data={byTipo} />
      case 'byArea': return <PieCard key={id} title="Por área de trabajo" icon="ti-building" data={byArea} colors={[C[1],C[2],C[3],C[4],C[5]]} />
      case 'byGrupo': return byGrupo.length > 0 ? <HBarCard key={id} full title="Personal por grupo de trabajo" icon="ti-hierarchy" data={byGrupo} xKey="label" bars={[{ key: 'count', label: 'Total' }, { key: 'activos', label: 'Activos' }]} colors={[C[0], C[1]]} /> : null
      case 'timeline': return timeline.length > 0 ? <TimelineCard key={id} data={timeline} /> : null
      case 'modHoras': return moduloStats.length > 0 ? <VBarCard key={id} full title="Horas dictadas por módulo" icon="ti-clock-hour-3" data={moduloStats} xKey="modulo" bars={[{ key: 'horas', label: 'Horas' }]} colors={[C[0]]} /> : null
      case 'modPers': return moduloStats.length > 0 ? <VBarCard key={id} full title="Participantes únicos por módulo" icon="ti-users" data={moduloStats} xKey="modulo" bars={[{ key: 'personas', label: 'Personas' }]} colors={[C[1]]} /> : null
      case 'tipoAct': return <PieCard key={id} title="Tipo de actividad registrada" icon="ti-list-check" data={tipoAct} />
      case 'notasMod': return moduloStats.filter(m => m.nota).length > 0 ? <VBarCard key={id} full title="Nota promedio por módulo" icon="ti-star" data={moduloStats.filter(m => m.nota)} xKey="modulo" bars={[{ key: 'nota', label: 'Nota promedio' }]} colors={[C[2]]} /> : null
      case 'distNotas': return <DistNotasCard key={id} data={distNotas} />
      case 'cumplim': return cumplimiento.length > 0 ? <CumplimientoCard key={id} data={cumplimiento} /> : null
      case 'topBV': return topBV.length > 0 ? <RankingCard key={id} icon="ti-award" title="Mayor participación (horas)" data={topBV} high /> : null
      case 'botBV': return botBV.length > 0 ? <RankingCard key={id} icon="ti-alert-triangle" title="Menor participación (horas)" data={botBV} high={false} /> : null
      case 'instructores': return instructores.length > 0 ? <HBarCard key={id} full title="Sesiones por instructor" icon="ti-school" data={instructores} xKey="label" bars={[{ key: 'sesiones', label: 'Sesiones' }, { key: 'horas', label: 'Horas' }]} colors={[C[0], C[2]]} /> : null
      case 'empresas': return empresas.length > 0 ? <HBarCard key={id} full title="Empresa instructora" icon="ti-building-factory-2" data={empresas} xKey="label" bars={[{ key: 'sesiones', label: 'Sesiones' }]} colors={[C[4]]} /> : null
      default: return null
    }
  }

  const now = new Date().toLocaleDateString('es-BO', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })

  return (
    <div style={{ display: 'flex', minHeight: '100vh' }}>

      {/* ── Sidebar ── */}
      <aside className="no-print" style={{
        width: 256, minWidth: 256, background: '#fff',
        borderRight: '1px solid #e5e7eb',
        display: 'flex', flexDirection: 'column', gap: 0,
        position: 'sticky', top: 0, height: '100vh', overflowY: 'auto',
      }}>
        {/* Logo */}
        <div style={{ padding: '1.25rem 1rem 1rem', borderBottom: '1px solid #f3f4f6' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 9, marginBottom: 3 }}>
            <div style={{ width: 30, height: 30, borderRadius: 8, background: '#1d4ed8', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <i className="ti ti-shield-half" style={{ fontSize: 16, color: '#fff' }} aria-hidden="true" />
            </div>
            <div>
              <p style={{ fontSize: 13, fontWeight: 600, margin: 0 }}>Reportes</p>
              <p style={{ fontSize: 10, color: '#9ca3af', margin: 0 }}>INSEIN – Brigadas de emergencia</p>
            </div>
          </div>
        </div>

        {/* Sección 1: Cargar */}
        <div style={{ padding: '1rem', borderBottom: '1px solid #f3f4f6' }}>
          <p style={{ fontSize: 10, fontWeight: 600, color: '#9ca3af', textTransform: 'uppercase', letterSpacing: '0.06em', margin: '0 0 8px' }}>1 · Cargar datos</p>
          <UploadZone onFile={handleFile} loading={loading} error={error} fileInfo={fileInfo} dragging={dragging} setDragging={setDragging} />
        </div>

        {stats && (
          <>
            {/* Sección 2: Título */}
            <div style={{ padding: '1rem', borderBottom: '1px solid #f3f4f6' }}>
              <p style={{ fontSize: 10, fontWeight: 600, color: '#9ca3af', textTransform: 'uppercase', letterSpacing: '0.06em', margin: '0 0 8px' }}>2 · Título del reporte</p>
              <input type="text" value={title} onChange={e => setTitle(e.target.value)} style={{ width: '100%' }} placeholder="Título del reporte..." />
            </div>

            {/* Sección 3: Indicadores */}
            <div style={{ padding: '1rem', flex: 1 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                <p style={{ fontSize: 10, fontWeight: 600, color: '#9ca3af', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0 }}>3 · Indicadores</p>
                <div style={{ display: 'flex', gap: 4 }}>
                  <button style={{ fontSize: 10, padding: '2px 7px' }} onClick={() => setSelected(INDICATORS.map(i => i.id))}>Todos</button>
                  <button style={{ fontSize: 10, padding: '2px 7px' }} onClick={() => setSelected([])}>Ninguno</button>
                </div>
              </div>
              {groups.map(g => (
                <div key={g} style={{ marginBottom: 12 }}>
                  <p style={{ fontSize: 9, fontWeight: 600, color: '#d1d5db', textTransform: 'uppercase', letterSpacing: '0.07em', margin: '0 0 5px' }}>{g}</p>
                  {INDICATORS.filter(i => i.group === g).map(ind => (
                    <label key={ind.id} style={{ display: 'flex', alignItems: 'center', gap: 7, padding: '3px 0', cursor: 'pointer', fontSize: 12, color: selected.includes(ind.id) ? '#111827' : '#9ca3af', transition: 'color 0.1s' }}>
                      <input type="checkbox" checked={selected.includes(ind.id)} onChange={() => toggle(ind.id)} style={{ accentColor: '#1d4ed8', width: 13, height: 13 }} />
                      <i className={`ti ${ind.icon}`} style={{ fontSize: 13 }} aria-hidden="true" />
                      {ind.label}
                    </label>
                  ))}
                </div>
              ))}
            </div>

            {/* Botón PDF */}
            <div style={{ padding: '1rem', borderTop: '1px solid #f3f4f6' }}>
              <button className="primary" style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, padding: '0.625rem' }} onClick={() => window.print()}>
                <i className="ti ti-file-type-pdf" aria-hidden="true" />
                Exportar PDF
              </button>
            </div>
          </>
        )}
      </aside>

      {/* ── Área principal ── */}
      <main className="print-area" style={{ flex: 1, padding: '2rem', minWidth: 0, background: '#f8f9fa' }}>
        {!stats ? (
          /* Estado vacío */
          <div style={{ maxWidth: 600, margin: '4rem auto', textAlign: 'center' }}>
            <div style={{ width: 72, height: 72, borderRadius: 20, background: '#eff6ff', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 1.5rem' }}>
              <i className="ti ti-file-spreadsheet" style={{ fontSize: 36, color: '#1d4ed8' }} aria-hidden="true" />
            </div>
            <h1 style={{ fontSize: 20, fontWeight: 600, marginBottom: 8 }}>Sube tu archivo para comenzar</h1>
            <p style={{ color: '#6b7280', marginBottom: 24, lineHeight: 1.7 }}>
              Carga <strong>DB_CAPACITACION.xlsx</strong> desde el panel izquierdo. El sistema analiza automáticamente las hojas <code>BD_personal</code>, <code>BD_seg</code> y <code>_tbl2</code>.
            </p>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, textAlign: 'left' }}>
              {[
                { icon: 'ti-users', t: 'Personal', d: 'Activos, bajas, tipo de brigada, área y grupo' },
                { icon: 'ti-clock', t: 'Horas', d: 'Horas dictadas por módulo e instructor' },
                { icon: 'ti-star', t: 'Evaluaciones', d: 'Notas promedio y distribución por módulo' },
                { icon: 'ti-circle-check', t: 'Cumplimiento', d: '% del programa completado por brigadista' },
              ].map(({ icon, t, d }) => (
                <div key={t} style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 10, padding: '0.875rem 1rem' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 4 }}>
                    <i className={`ti ${icon}`} style={{ fontSize: 14, color: '#1d4ed8' }} aria-hidden="true" />
                    <span style={{ fontSize: 13, fontWeight: 500 }}>{t}</span>
                  </div>
                  <p style={{ fontSize: 11, color: '#9ca3af', margin: 0, lineHeight: 1.5 }}>{d}</p>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <>
            {/* Header del reporte */}
            <div className="report-header" style={{ marginBottom: 24, paddingBottom: 20, borderBottom: '1px solid #e5e7eb' }}>
              <h1 style={{ fontSize: 22, fontWeight: 600, marginBottom: 4 }}>{title}</h1>
              <div style={{ display: 'flex', alignItems: 'center', gap: 16, fontSize: 12, color: '#9ca3af' }}>
                <span><i className="ti ti-calendar" style={{ marginRight: 4 }} aria-hidden="true" />{now}</span>
                {fileInfo && <span style={{ padding: '2px 10px', background: '#eff6ff', borderRadius: 20, color: '#1d4ed8', fontWeight: 500 }}>{fileInfo.name} · {fileInfo.size}</span>}
              </div>
            </div>

            {selected.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '3rem', color: '#9ca3af' }}>
                <i className="ti ti-checks" style={{ fontSize: 36, display: 'block', marginBottom: 8, opacity: 0.4 }} aria-hidden="true" />
                <p>Selecciona indicadores en el panel izquierdo</p>
              </div>
            ) : (
              <div className="indicator-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 16 }}>
                {INDICATORS.filter(i => selected.includes(i.id)).map(i => render(i.id))}
              </div>
            )}
          </>
        )}
      </main>
    </div>
  )
}
