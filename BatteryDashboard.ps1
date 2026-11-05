<#
.SYNOPSIS
  Battery Dashboard - reads Windows battery data and opens a clean HTML dashboard.

.DESCRIPTION
  Uses `powercfg /batteryreport` plus live WMI readings to show battery health,
  cycle count, current charge, charge history and daily usage.
  Output is saved to %LOCALAPPDATA%\BatteryDashboard\dashboard.html.

.PARAMETER NoOpen
  Generate the dashboard without opening it in the browser.

.PARAMETER CreateShortcut
  Also put a "Battery Dashboard" shortcut on the Desktop.

.EXAMPLE
  .\BatteryDashboard.ps1
  .\BatteryDashboard.ps1 -CreateShortcut
#>
param([switch]$NoOpen, [switch]$CreateShortcut)

$ErrorActionPreference = 'Stop'

$outDir   = Join-Path $env:LOCALAPPDATA 'BatteryDashboard'
$xmlPath  = Join-Path $outDir 'battery-report.xml'
$fullPath = Join-Path $outDir 'battery-report.html'
$htmlPath = Join-Path $outDir 'dashboard.html'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

if ($CreateShortcut) {
    $lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Battery Dashboard.lnk'
    $s = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
    $s.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $s.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    $s.WorkingDirectory = Split-Path $PSCommandPath
    $icon = "$env:SystemRoot\System32\batmeter.dll"
    if (Test-Path $icon) { $s.IconLocation = "$icon,0" }
    $s.Save()
    Write-Host "Shortcut created: $lnk"
}

Write-Host 'Reading battery data...'
powercfg /batteryreport /xml /output "$xmlPath" | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'powercfg could not create a battery report. Does this PC have a battery?' }
powercfg /batteryreport /output "$fullPath" | Out-Null

[xml]$x = Get-Content -Raw -Encoding UTF8 $xmlPath
$r = $x.BatteryReport

function Sec($d) {
    if ([string]::IsNullOrWhiteSpace($d)) { return 0 }
    try { [int][Xml.XmlConvert]::ToTimeSpan($d).TotalSeconds } catch { 0 }
}

# --- Battery (sum capacities if there is more than one pack) ---
$packs = @($r.Batteries.Battery)
if ($packs.Count -eq 0) { throw 'No battery found in the report.' }
$battery = [ordered]@{
    id           = ($packs | ForEach-Object { $_.Id }) -join ' + '
    manufacturer = $packs[0].Manufacturer
    chemistry    = $packs[0].Chemistry
    design       = ($packs | ForEach-Object { [int]$_.DesignCapacity } | Measure-Object -Sum).Sum
    full         = ($packs | ForEach-Object { [int]$_.FullChargeCapacity } | Measure-Object -Sum).Sum
    cycles       = ($packs | ForEach-Object { [int]$_.CycleCount } | Measure-Object -Maximum).Maximum
    relative     = ($packs[0].RelativeCapacity -eq '1')
}

# --- Live readings ---
$live = [ordered]@{ percent = $null; onAc = $null; charging = $false; rateW = $null; remaining = $null }
try {
    $wb = Get-CimInstance Win32_Battery | Select-Object -First 1
    if ($wb) { $live.percent = [int]$wb.EstimatedChargeRemaining }
} catch {}
try {
    $bs = Get-CimInstance -Namespace root\wmi -ClassName BatteryStatus | Select-Object -First 1
    if ($bs) {
        $live.onAc      = [bool]$bs.PowerOnline
        $live.charging  = [bool]$bs.Charging
        $live.remaining = [int]$bs.RemainingCapacity
        if ($bs.Charging -and $bs.ChargeRate -gt 0)          { $live.rateW = [math]::Round($bs.ChargeRate / 1000, 1) }
        elseif ($bs.Discharging -and $bs.DischargeRate -gt 0) { $live.rateW = -[math]::Round($bs.DischargeRate / 1000, 1) }
    }
} catch {}

# --- Recent usage (charge level over the last 7 days) ---
$usage = foreach ($u in @($r.RecentUsage.UsageEntry)) {
    $fc = [double]$u.FullChargeCapacity
    if ($fc -le 0) { continue }
    [ordered]@{
        t    = $u.LocalTimestamp
        pct  = [math]::Round([math]::Min(100, 100 * [double]$u.ChargeCapacity / $fc), 1)
        ac   = ($u.Ac -eq '1')
        type = $u.EntryType
    }
}

# --- Capacity and usage history (one entry per day or week) ---
$history = foreach ($h in @($r.History.HistoryEntry)) {
    [ordered]@{
        start    = $h.LocalStartDate
        end      = $h.LocalEndDate
        design   = [int]$h.DesignCapacity
        full     = [int]$h.FullChargeCapacity
        cycles   = [int]$h.CycleCount
        acActive = Sec $h.ActiveAcTime
        dcActive = Sec $h.ActiveDcTime
        csAc     = Sec $h.CsAcTime
        csDc     = Sec $h.CsDcTime
    }
}

$sys = $r.SystemInformation
$data = [ordered]@{
    generated = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    computer  = $sys.ComputerName
    model     = $sys.SystemProductName
    bios      = $sys.BIOSVersion
    os        = $sys.OSBuild
    battery   = $battery
    live      = $live
    runtime   = [ordered]@{
        full   = Sec $r.RuntimeEstimates.FullChargeCapacity.ActiveRuntime
        design = Sec $r.RuntimeEstimates.DesignCapacity.ActiveRuntime
    }
    usage     = @($usage)
    history   = @($history)
    fullReport = ([Uri]$fullPath).AbsoluteUri
}
$json = ($data | ConvertTo-Json -Depth 6 -Compress) -replace '</', '<\/'

$template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Battery Dashboard</title>
<style>
:root{
  --bg:#f5f6f8;--card:#ffffff;--text:#16181d;--muted:#6b7280;--grid:#e6e8ec;--border:#e6e8ec;
  --accent:#2563eb;--accent-soft:rgba(37,99,235,.12);--ac:rgba(22,163,74,.10);
  --good:#16a34a;--warn:#d97706;--bad:#dc2626;--bar2:#cbd5e1;
}
@media (prefers-color-scheme:dark){:root{
  --bg:#0e1014;--card:#171a20;--text:#e8eaef;--muted:#9aa1ad;--grid:#262a32;--border:#242830;
  --accent:#60a5fa;--accent-soft:rgba(96,165,250,.16);--ac:rgba(74,222,128,.09);
  --good:#4ade80;--warn:#fbbf24;--bad:#f87171;--bar2:#3b4352;
}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);
  font:14px/1.45 "Segoe UI Variable Text","Segoe UI",system-ui,sans-serif}
.wrap{max-width:1100px;margin:0 auto;padding:28px 20px 40px}
header{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;margin-bottom:20px}
h1{font-size:26px;margin:0;font-weight:650;letter-spacing:-.01em}
.sub{color:var(--muted);margin-top:2px}
.stamp{color:var(--muted);font-size:13px;text-align:right}
.stamp a{color:var(--accent);text-decoration:none}
.grid{display:grid;gap:14px}
.tiles{grid-template-columns:repeat(auto-fit,minmax(180px,1fr));margin-bottom:14px}
.card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:18px}
.label{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.06em;font-weight:600}
.big{font-size:34px;font-weight:650;letter-spacing:-.02em;margin-top:6px;font-variant-numeric:tabular-nums}
.big small{font-size:15px;font-weight:500;color:var(--muted);margin-left:4px}
.note{color:var(--muted);font-size:13px;margin-top:4px}
.health{display:flex;align-items:center;gap:16px}
.pill{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:600;margin-top:6px}
.pill.good{color:var(--good);background:color-mix(in srgb,var(--good) 14%,transparent)}
.pill.warn{color:var(--warn);background:color-mix(in srgb,var(--warn) 14%,transparent)}
.pill.bad{color:var(--bad);background:color-mix(in srgb,var(--bad) 14%,transparent)}
.meter{height:8px;border-radius:99px;background:var(--grid);margin-top:12px;overflow:hidden}
.meter i{display:block;height:100%;border-radius:99px;background:var(--accent)}
.chart-card{margin-bottom:14px}
.chart-head{display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap;margin-bottom:8px}
.chart-head h2{font-size:15px;margin:0;font-weight:600}
.legend{display:flex;gap:14px;color:var(--muted);font-size:12px}
.legend span::before{content:"";display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:-1px;background:var(--c)}
.chart{width:100%;min-height:60px}
.chart svg{display:block;overflow:visible}
.chart text{fill:var(--muted);font-size:11px}
.empty{color:var(--muted);padding:24px 0;text-align:center}
.two{grid-template-columns:repeat(auto-fit,minmax(320px,1fr))}
table{width:100%;border-collapse:collapse}
td{padding:7px 0;border-bottom:1px solid var(--border);vertical-align:top}
tr:last-child td{border-bottom:0}
td:first-child{color:var(--muted);width:44%}
ul.tips{margin:6px 0 0;padding-left:18px;color:var(--muted)}
ul.tips li{margin:4px 0}
#tip{position:fixed;pointer-events:none;background:var(--text);color:var(--bg);padding:6px 9px;border-radius:8px;
  font-size:12px;line-height:1.35;opacity:0;transition:opacity .1s;white-space:nowrap;z-index:9}
@media (max-width:520px){.wrap{padding:20px 16px}.big{font-size:28px}.stamp{text-align:left}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div><h1>Battery Health</h1><div class="sub" id="model"></div></div>
    <div class="stamp"><div id="stamp"></div><a id="full" href="#">Open full Windows report &rarr;</a></div>
  </header>

  <section class="grid tiles">
    <div class="card health">
      <svg width="92" height="92" viewBox="0 0 92 92" id="ring"></svg>
      <div><div class="label">Health</div><div class="big" id="hp"></div><span class="pill" id="verdict"></span></div>
    </div>
    <div class="card"><div class="label">Current charge</div><div class="big" id="pct"></div>
      <div class="note" id="state"></div><div class="meter"><i id="pctbar"></i></div></div>
    <div class="card"><div class="label">Charge cycles</div><div class="big" id="cycles"></div><div class="note" id="cyclenote"></div></div>
    <div class="card"><div class="label">Capacity</div><div class="big" id="cap"></div><div class="note" id="capnote"></div></div>
    <div class="card"><div class="label">Est. runtime</div><div class="big" id="rt"></div><div class="note">At full charge (Windows estimate)</div></div>
  </section>

  <section class="card chart-card">
    <div class="chart-head"><h2>Charge level &middot; last 7 days</h2>
      <div class="legend"><span style="--c:var(--accent)">Charge %</span><span style="--c:color-mix(in srgb,var(--good) 35%,transparent)">Plugged in</span></div></div>
    <div class="chart" id="charge"></div>
  </section>

  <section class="grid two">
    <div class="card chart-card">
      <div class="chart-head"><h2>Health over time</h2><div class="legend"><span style="--c:var(--good)">Full charge / design</span></div></div>
      <div class="chart" id="health"></div>
    </div>
    <div class="card chart-card">
      <div class="chart-head"><h2>Screen-on time per day</h2>
        <div class="legend"><span style="--c:var(--accent)">On battery</span><span style="--c:var(--bar2)">Plugged in</span></div></div>
      <div class="chart" id="bars"></div>
    </div>
  </section>

  <section class="grid two">
    <div class="card"><div class="label" style="margin-bottom:6px">Details</div><table id="details"></table></div>
    <div class="card"><div class="label">Tips</div><ul class="tips" id="tips"></ul></div>
  </section>
</div>
<div id="tip"></div>

<script>
const D = __DATA__;
const NS = 'http://www.w3.org/2000/svg';
const $ = s => document.querySelector(s);
const el = (tag, attrs, parent) => { const e = document.createElementNS(NS, tag);
  for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; };
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const dur = s => { if (!s) return '0m'; const h = Math.floor(s/3600), m = Math.round((s%3600)/60); return h ? `${h}h ${m}m` : `${m}m`; };
const wh = mwh => (mwh/1000).toFixed(1);
const dShort = d => d.toLocaleDateString(undefined, {day:'numeric', month:'short'});
const dTime = d => d.toLocaleString(undefined, {weekday:'short', day:'numeric', month:'short', hour:'numeric', minute:'2-digit'});

const tip = $('#tip');
function showTip(html, ev){ tip.innerHTML = html; tip.style.opacity = 1;
  const w = tip.offsetWidth, x = ev.clientX + 14 + w > innerWidth ? ev.clientX - w - 14 : ev.clientX + 14;
  tip.style.left = x + 'px'; tip.style.top = (ev.clientY - 40) + 'px'; }
const hideTip = () => tip.style.opacity = 0;

// ---------- Tiles ----------
const b = D.battery, L = D.live;
const health = b.design > 0 && !b.relative ? b.full / b.design * 100 : null;
const v = health == null ? ['Unknown','warn'] : health >= 80 ? ['Good','good'] : health >= 60 ? ['Fair','warn'] : ['Consider replacing','bad'];
$('#model').textContent = [D.model, D.computer].filter(Boolean).join(' · ');
$('#stamp').textContent = 'Updated ' + dTime(new Date(D.generated));
$('#full').href = D.fullReport;
$('#hp').innerHTML = health == null ? '—' : `${Math.min(health,100).toFixed(1)}<small>%</small>`;
$('#verdict').textContent = v[0]; $('#verdict').className = 'pill ' + v[1];
(() => { const r = 38, c = 2*Math.PI*r, svg = $('#ring'), p = Math.max(0, Math.min(health ?? 0, 100)) / 100;
  el('circle', {cx:46, cy:46, r, fill:'none', stroke:css('--grid'), 'stroke-width':9}, svg);
  el('circle', {cx:46, cy:46, r, fill:'none', stroke:css('--' + v[1]), 'stroke-width':9, 'stroke-linecap':'round',
    'stroke-dasharray':`${c*p} ${c}`, transform:'rotate(-90 46 46)'}, svg); })();

const pct = L.percent ?? (D.usage.length ? Math.round(D.usage.at(-1).pct) : null);
$('#pct').innerHTML = pct == null ? '—' : `${pct}<small>%</small>`;
$('#pctbar').style.width = (pct ?? 0) + '%';
let state = L.charging ? 'Charging' : L.onAc ? 'Plugged in' : L.onAc === false ? 'On battery' : '';
if (L.rateW) state += ` · ${Math.abs(L.rateW)} W`;
if (L.rateW < 0 && L.remaining) state += ` · ~${dur(L.remaining / (-L.rateW*1000) * 3600)} left`;
$('#state').textContent = state || 'Live status unavailable';

$('#cycles').textContent = b.cycles ?? '—';
$('#cyclenote').textContent = 'Most laptop batteries are rated for 500–1000';
$('#cap').innerHTML = `${wh(b.full)}<small>/ ${wh(b.design)} Wh</small>`;
$('#capnote').textContent = health == null ? 'Battery reports relative capacity only'
  : health >= 100 ? 'No wear yet' : `${(100 - health).toFixed(1)}% wear`;
$('#rt').textContent = D.runtime.full ? dur(D.runtime.full) : '—';

// ---------- Details & tips ----------
[['Battery', b.id], ['Manufacturer', b.manufacturer], ['Chemistry', b.chemistry],
 ['Design capacity', b.design.toLocaleString() + ' mWh'], ['Full charge capacity', b.full.toLocaleString() + ' mWh'],
 ['Cycle count', b.cycles], ['Model', D.model], ['BIOS', D.bios], ['OS build', D.os]]
 .forEach(([k, val]) => { const tr = document.createElement('tr');
   tr.innerHTML = `<td>${k}</td><td>${val ?? '—'}</td>`; $('#details').appendChild(tr); });
const tips = [
  'Health is full-charge capacity divided by design capacity. Under 80% is the usual time to replace.',
  'If the laptop is mostly plugged in, cap charging at 80% (MyASUS → Battery Care Mode on ASUS laptops).',
  'Heat wears batteries fastest: keep vents clear and avoid charging on soft surfaces.',
  'Run this dashboard every few months to watch the trend.'];
tips.forEach(t => { const li = document.createElement('li'); li.textContent = t; $('#tips').appendChild(li); });

// ---------- Charts ----------
function frame(box, H, m){ box.innerHTML = ''; const W = box.clientWidth;
  const svg = el('svg', {viewBox:`0 0 ${W} ${H}`, width:W, height:H}, box);
  return {svg, W, H, iw: W - m.l - m.r, ih: H - m.t - m.b}; }
function hgrid(f, m, ticks, y, fmt){ ticks.forEach(t => {
  el('line', {x1:m.l, x2:f.W - m.r, y1:y(t), y2:y(t), stroke:css('--grid')}, f.svg);
  el('text', {x:m.l - 8, y:y(t) + 4, 'text-anchor':'end'}, f.svg).textContent = fmt(t); }); }

function chargeChart(){
  const box = $('#charge');
  const pts = D.usage.map(u => ({t:new Date(u.t).getTime(), v:u.pct, ac:u.ac}))
    .filter(p => !isNaN(p.t)).sort((a, z) => a.t - z.t);
  if (pts.length < 2) { box.innerHTML = '<p class="empty">Not enough recent data yet. Use the laptop for a day and run this again.</p>'; return; }
  const m = {l:40, r:12, t:10, b:26}, f = frame(box, 230, m);
  const t0 = pts[0].t, t1 = Math.max(pts.at(-1).t, new Date(D.generated).getTime());
  const x = t => m.l + (t - t0) / (t1 - t0) * f.iw, y = v => m.t + (1 - v/100) * f.ih;
  pts.forEach((p, i) => { if (!p.ac) return; const x2 = x(pts[i+1]?.t ?? t1);
    el('rect', {x:x(p.t), y:m.t, width:Math.max(0, x2 - x(p.t)), height:f.ih, fill:css('--ac')}, f.svg); });
  hgrid(f, m, [0, 25, 50, 75, 100], y, t => t + '%');
  const day = new Date(t0); day.setHours(24, 0, 0, 0);
  for (let d = day.getTime(); d < t1; d += 864e5) {
    el('line', {x1:x(d), x2:x(d), y1:m.t, y2:m.t + f.ih, stroke:css('--grid'), 'stroke-dasharray':'2 3'}, f.svg);
    el('text', {x:x(d) + 4, y:f.H - 8}, f.svg).textContent = new Date(d).toLocaleDateString(undefined, {weekday:'short', day:'numeric'});
  }
  const line = pts.map((p, i) => `${i ? 'L' : 'M'}${x(p.t).toFixed(1)},${y(p.v).toFixed(1)}`).join('') + `L${x(t1)},${y(pts.at(-1).v)}`;
  el('path', {d:`${line}L${x(t1)},${y(0)}L${x(t0)},${y(0)}Z`, fill:css('--accent-soft')}, f.svg);
  el('path', {d:line, fill:'none', stroke:css('--accent'), 'stroke-width':2, 'stroke-linejoin':'round'}, f.svg);
  const cross = el('line', {y1:m.t, y2:m.t + f.ih, stroke:css('--muted'), opacity:0}, f.svg);
  const dot = el('circle', {r:4.5, fill:css('--accent'), stroke:css('--card'), 'stroke-width':2, opacity:0}, f.svg);
  const hit = el('rect', {x:m.l, y:m.t, width:f.iw, height:f.ih, fill:'transparent'}, f.svg);
  hit.addEventListener('mousemove', ev => {
    const t = t0 + (ev.offsetX - m.l) / f.iw * (t1 - t0);
    let p = pts[0]; for (const q of pts) { if (q.t <= t) p = q; else break; }
    cross.setAttribute('x1', ev.offsetX); cross.setAttribute('x2', ev.offsetX); cross.setAttribute('opacity', .4);
    dot.setAttribute('cx', ev.offsetX); dot.setAttribute('cy', y(p.v)); dot.setAttribute('opacity', 1);
    showTip(`<b>${Math.round(p.v)}%</b> · ${p.ac ? 'Plugged in' : 'On battery'}<br>${dTime(new Date(t))}`, ev);
  });
  hit.addEventListener('mouseleave', () => { hideTip(); cross.setAttribute('opacity', 0); dot.setAttribute('opacity', 0); });
}

function healthChart(){
  const box = $('#health');
  const pts = D.history.filter(h => h.design > 0 && h.full > 0)
    .map(h => ({t:new Date(h.start).getTime(), end:new Date(h.end), v:h.full / h.design * 100, cycles:h.cycles, full:h.full}));
  if (!pts.length || b.relative) { box.innerHTML = '<p class="empty">No capacity history available.</p>'; return; }
  const m = {l:44, r:14, t:12, b:26}, f = frame(box, 210, m);
  const vals = pts.map(p => p.v), lo = Math.min(75, Math.floor(Math.min(...vals) - 5)), hi = Math.max(105, Math.ceil(Math.max(...vals) + 2));
  const t0 = pts[0].t, t1 = pts.at(-1).t, span = t1 - t0 || 1;
  const x = t => pts.length === 1 ? m.l + f.iw/2 : m.l + (t - t0) / span * f.iw, y = v => m.t + (hi - v) / (hi - lo) * f.ih;
  const step = (hi - lo) > 30 ? 10 : 5, ticks = []; for (let t = Math.ceil(lo/step)*step; t <= hi; t += step) ticks.push(t);
  hgrid(f, m, ticks, y, t => t + '%');
  el('line', {x1:m.l, x2:f.W - m.r, y1:y(80), y2:y(80), stroke:css('--bad'), 'stroke-dasharray':'5 4', opacity:.7}, f.svg);
  el('text', {x:f.W - m.r, y:y(80) - 5, 'text-anchor':'end', style:`fill:${css('--bad')}`}, f.svg).textContent = '80% replace threshold';
  const n = Math.min(pts.length, Math.max(2, Math.floor(f.iw / 90)));
  for (let i = 0; i < n; i++) { const p = pts[Math.round(i * (pts.length - 1) / Math.max(1, n - 1))];
    el('text', {x:x(p.t), y:f.H - 8, 'text-anchor': pts.length === 1 ? 'middle' : i === 0 ? 'start' : i === n - 1 ? 'end' : 'middle'}, f.svg).textContent = dShort(new Date(p.t)); }
  if (pts.length > 1) el('path', {d:pts.map((p, i) => `${i ? 'L' : 'M'}${x(p.t)},${y(p.v)}`).join(''), fill:'none', stroke:css('--good'), 'stroke-width':2}, f.svg);
  pts.forEach(p => { const c = el('circle', {cx:x(p.t), cy:y(p.v), r:pts.length > 40 ? 2.5 : 4, fill:css('--good'), stroke:css('--card'), 'stroke-width':1.5}, f.svg);
    c.addEventListener('mousemove', ev => showTip(`<b>${p.v.toFixed(1)}%</b> health · ${p.full.toLocaleString()} mWh<br>${dShort(new Date(p.t))}${p.cycles ? ' · ' + p.cycles + ' cycles' : ''}`, ev));
    c.addEventListener('mouseleave', hideTip); });
}

function barsChart(){
  const box = $('#bars');
  const days = D.history.slice(-14).map(h => ({d:new Date(h.start), bat:h.dcActive/3600, ac:h.acActive/3600}));
  if (!days.length) { box.innerHTML = '<p class="empty">No usage history yet.</p>'; return; }
  const m = {l:36, r:8, t:12, b:26}, f = frame(box, 210, m);
  const max = Math.max(1, Math.ceil(Math.max(...days.map(d => d.bat + d.ac))));
  const y = v => m.t + (1 - v / max) * f.ih, stepH = max > 8 ? Math.ceil(max / 4) : max > 4 ? 2 : 1, ticks = [];
  for (let t = 0; t <= max; t += stepH) ticks.push(t);
  hgrid(f, m, ticks, y, t => t + 'h');
  const bw = f.iw / days.length, w = Math.min(28, bw * .62), every = Math.ceil(days.length / Math.max(1, Math.floor(f.iw / 46)));
  days.forEach((d, i) => { const cx = m.l + bw * i + bw / 2, g = el('g', {}, f.svg);
    el('rect', {x:cx - bw/2, y:m.t, width:bw, height:f.ih, fill:'transparent'}, g);
    if (d.ac > 0) el('rect', {x:cx - w/2, y:y(d.bat + d.ac), width:w, height:Math.max(0, y(d.bat) - y(d.bat + d.ac)), rx:3, fill:css('--bar2')}, g);
    if (d.bat > 0) el('rect', {x:cx - w/2, y:y(d.bat), width:w, height:Math.max(0, y(0) - y(d.bat)), rx:3, fill:css('--accent')}, g);
    if (i % every === 0) el('text', {x:cx, y:f.H - 8, 'text-anchor':'middle'}, f.svg).textContent = dShort(d.d);
    g.addEventListener('mousemove', ev => showTip(`<b>${dShort(d.d)}</b><br>On battery: ${dur(d.bat*3600)}<br>Plugged in: ${dur(d.ac*3600)}`, ev));
    g.addEventListener('mouseleave', hideTip); });
}

function render(){ chargeChart(); healthChart(); barsChart(); }
render();
let rt; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(render, 120); });
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => location.reload());
</script>
</body>
</html>
'@

$html = $template.Replace('__DATA__', $json)
[IO.File]::WriteAllText($htmlPath, $html, (New-Object Text.UTF8Encoding $false))
Write-Host "Dashboard saved: $htmlPath"
if (-not $NoOpen) { Start-Process $htmlPath }
