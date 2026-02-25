/* Variance Report – NOI-first, bucket-aware, T12 order, Stoa-ish theme, UX upgrades */

/* =========================
   1) Aliases & columns
   ========================= */
// Dataset alias: t12vsbudget (Domo dataset name) or varianceReport (manifest default)
const DATASET_ALIASES = ['t12vsbudget', 'varianceReport'];

const COLS = {
  prop: 'Property',
  parentCat: 'ParentCategoryName',
  cat: 'CategoryName',
  glNum: 'AccountNumber',
  glName: 'Account',
  bud: 'BudgetValue',
  t12: 'T12Value',            // UI label → “Actual”
  date: 'Date',
  units: 'Units',
  birth: 'BirthOrder',
  city: 'City',
  state: 'State',
  status: 'Status',
  finStatus: 'FinancingStatus',
  productType: 'ProductType',
  fullAddress: 'FullAddress',
  region: 'Region',
  construction: 'ConstructionStatus'
};

const LEVELS = [
  { key: COLS.prop, label: 'Property' },
  { key: '__bucket', label: 'Bucket' },
  { key: COLS.parentCat, label: 'Parent Category' },
  { key: COLS.cat, label: 'Category' },
  { key: COLS.glNum, label: 'GL Code',
    nameFn: r => `${r[COLS.glNum] ?? '—'}${r[COLS.glName] ? ' — ' + r[COLS.glName] : ''}` }
];

const STATE = { drill: [], rows: [], monthIndex: 0, months: [], selectedProperties: new Set() };
let _refreshPropertyListForMonth = null;
const $ = s => document.querySelector(s);

/* =========================
   2) Accounting buckets
   ========================= */
const ACCT_BUCKETS = {
  income: new Set([
    'total income','rental income','total rental income','net rental income','other income',
    'service related income','financial income','revenue- internal','revenue- retail'
  ]),
  opex: new Set([
    'total operating expenses','payroll & related','payroll and related','taxes & insurance','taxes and insurance','utilities',
    'maintenance & repairs','maintenance and repairs','administrative expenses','marketing expenses',
    'management fees','facility expenses','information technology expenses','professional expenses',
    'travel expenses','safety & training expenses','service related expenses','incentive expenses','business license expenses'
  ]),
  nonop: new Set([
    'non-operating expenses','other non-operating expenses','other non-operating','capital expenditures'
  ]),
  debt: new Set(['debt services','debt service'])
};

/* T12-ish ordering for buckets / parents / categories */
const ORDER_BUCKET = ['Income','Operating Expenses','Non-Operating','Debt Service','Other'];
const ORDER_PARENT = [
  'Rental Income','Net Rental Income','Other Income','Total Income',
  'Payroll & Related','Taxes & Insurance','Utilities','Maintenance & Repairs',
  'Administrative Expenses','Marketing Expenses','Management Fees','Total Operating Expenses',
  'Other Operating Expenses',
  'Other Non-Operating Expenses','Non-Operating Expenses',
  'Capital Expenditures',
  'Debt Services','Debt Service'
].map(s => s.toLowerCase());

/* Categories where a positive variance is GOOD even in expense buckets */
const SPECIAL_POSITIVE_GOOD = new Set([
  'gain to lease'
]);

/* Loss/reduction categories: more negative actual = worse. Negative variance = BAD (red), positive = GOOD (green). */
const NEGATIVE_VARIANCE_IS_BAD = new Set([
  'vacancy loss', 'loss to lease', 'concession loss', 'bad debt', 'non-revenue wash',
  'vacancy and credit loss', 'loss to lease and vacancy', 'economic vacancy loss',
  'upfront concessions'
]);

/* Bottom-line / profit metrics: positive variance = good, negative = bad (same as NOI). Overrides bucket. */
const BOTTOM_LINE_POSITIVE_GOOD = new Set([
  'net operating income (loss)', 'net income (loss)', 'net operating income', 'net income'
]);

/* =========================
   3) Utilities
   ========================= */
function fmtMoney(n){ if(n==null || Number.isNaN(n)) return '—'; return n.toLocaleString(undefined,{style:'currency',currency:'USD',maximumFractionDigits:0}); }
function fmtPct(n){ if(n==null || Number.isNaN(n)) return '—'; return n.toLocaleString(undefined,{maximumFractionDigits:2}) + '%'; }
function toNum(v){
  if(v == null) return null;
  let s = String(v).trim();
  const neg = /^\(.*\)$/.test(s);
  if (neg) s = s.slice(1,-1);
  s = s.replace(/[%$,]/g,'').replace(/\s+/g,'');
  const n = Number(s);
  return Number.isFinite(n) ? (neg ? -n : n) : null;
}
function _canon(s){ return (s||'').trim().toLowerCase(); }
function showError(err){
  const box = $('#error'), pre = $('#errorMsg');
  const msg = err?.message || (typeof err === 'string' ? err : JSON.stringify(err ?? {}, null, 2));
  if (pre) pre.textContent = msg || 'Unknown error';
  if (box) box.hidden = false;
  console.error('[t12vsbudget]', err);
}
function diag(s){
  const box = $('#error'), pre = $('#errorMsg');
  if (pre) pre.textContent = String(s);
  if (box) box.hidden = false;
}

/* Treat totals / nets as rollups so we don't double-count */
/* NOTE: "Net Operating Income (Loss)" is NOT a rollup - it's the actual NOI line item we want to use */
function isRollupRow(r){
  const p = _canon(r[COLS.parentCat]);
  const c = _canon(r[COLS.cat]);
  const hits = (s) =>
    s.startsWith('total ') ||
    s.includes('grand total') ||
    (s.includes('net operating income') && !s.includes('net operating income (loss)')) ||
    s === 'net rental income' || s === 'total rental income' ||
    s === 'total other income' || s === 'total income' ||
    s === 'total operating expenses' || (s.includes('noi') && !s.includes('net operating income (loss)'));
  return hits(p) || hits(c);
}

function bucketOf(r){
  const p = _canon(r[COLS.parentCat]);
  const c = _canon(r[COLS.cat]);
  const t = p || c;
  
  // Exclude NOI and Net Income (Loss) from Income bucket to prevent double counting
  // NOI is calculated as Income - Operating Expenses, so including it in Income would double count
  // Net Income (Loss) is a bottom-line item and should not be in expense buckets
  const isNOI = (p.includes('net operating income') || c.includes('net operating income')) || 
                 ((p.includes('noi') || c.includes('noi')) && !p.includes('noise') && !c.includes('noise'));
  const isNetIncome = (p.includes('net income') || c.includes('net income')) && 
                      !p.includes('net operating income') && !c.includes('net operating income');
  
  // If it's NOI or Net Income (Loss), exclude from all buckets (put in Other)
  if (isNOI || isNetIncome) return 'Other';
  
  // Check explicit bucket mappings first
  if (ACCT_BUCKETS.income.has(t)) return 'Income';
  if (ACCT_BUCKETS.opex.has(t))   return 'Operating Expenses';
  if (ACCT_BUCKETS.nonop.has(t))  return 'Non-Operating';
  if (ACCT_BUCKETS.debt.has(t))   return 'Debt Service';
  
  // For other income/revenue items (but not NOI or Net Income), assign to Income
  if (t.includes('income') || t.includes('revenue')) return 'Income';
  if (t.includes('debt'))                            return 'Debt Service';
  if (t.includes('non-oper'))                        return 'Non-Operating';
  if (t.includes('expense'))                         return 'Operating Expenses';
  
  return 'Other';
}
function isExpenseBucket(b){
  return b === 'Operating Expenses' || b === 'Non-Operating' || b === 'Debt Service' || b === 'Other';
}
function isGoodVariance(bucket, vDollar, parentName, catName){
  const n = _canon(catName) || _canon(parentName);
  const p = _canon(parentName);
  const round2 = (x) => (x != null && Number.isFinite(x)) ? Math.round(x * 100) / 100 : x;
  vDollar = round2(vDollar);
  // Bottom-line (NOI, Net Income): positive = good, negative = bad — never treat as expense.
  const isBottomLine = (s) => Array.from(BOTTOM_LINE_POSITIVE_GOOD).some(k => (s || '').includes(k));
  if (isBottomLine(n) || isBottomLine(p)) return vDollar >= 0;
  // Loss-type categories: more negative actual than budget = bad. So positive variance = good, negative = bad.
  const isLossCategory = (s) => Array.from(NEGATIVE_VARIANCE_IS_BAD).some(k => (s || '').includes(k));
  if (NEGATIVE_VARIANCE_IS_BAD.has(n) || NEGATIVE_VARIANCE_IS_BAD.has(p) || isLossCategory(n) || isLossCategory(p)) return vDollar >= 0;
  if (SPECIAL_POSITIVE_GOOD.has(n)) return vDollar >= 0; // Gain to Lease → more is good
  if (bucket === 'NOI' || bucket === 'Income') return vDollar >= 0;
  if (isExpenseBucket(bucket)) return vDollar <= 0;
  return vDollar >= 0;
}
function propertyMeta(rows){
  const r = rows.find(x=>x[COLS.prop]);
  return {
    units: r?.[COLS.units],
    city: r?.[COLS.city],
    state: r?.[COLS.state],
    status: r?.[COLS.status],
    finStatus: r?.[COLS.finStatus],
    region: r?.[COLS.region],
    productType: r?.[COLS.productType]
  };
}

/* =========================
   4) Data: fetch via alias
   ========================= */
async function fetchRowsByAlias(alias){
  const LIMIT = 50000;
  let offset = 0, all = [];

  // Omit fields param to avoid 400 from schema mismatches (dataset column names vary)
  const buildUrl = (off) => `/data/v2/${alias}?limit=${LIMIT}&offset=${off}`;

  while (true){
    const url = buildUrl(offset);
    let page;
    try { 
      page = await domo.get(url); 
    }
    catch (e) { 
      throw new Error(`Alias API failed for '${alias}' @ offset ${offset}. ${e?.message || JSON.stringify(e)}`); 
    }
    if (!Array.isArray(page)) throw new Error(`Unexpected alias response for '${alias}': ${JSON.stringify(page)}`);
    all = all.concat(page);
    if (page.length < LIMIT) break;
    offset += LIMIT;
  }

  // 🔎 Filter out Sold, Pre-Construction, and Under Contract properties
  all = all.filter(r => {
    const st = (r[COLS.status] || '').toLowerCase();
    return st !== 'sold' && st !== 'pre-construction' && st !== 'under contract';
  });

  for (const r of all){
    r[COLS.bud]  = toNum(r[COLS.bud]);
    r[COLS.t12]  = toNum(r[COLS.t12]);
    if (r[COLS.birth] != null) r[COLS.birth] = Number(String(r[COLS.birth]).replace(/[^\d.-]/g,'')) || 0;

    const v$ = (r[COLS.t12] ?? 0) - (r[COLS.bud] ?? 0);
    r.__var$ = Number.isFinite(v$) ? v$ : null;
    const base = r[COLS.bud];
    r.__varp = (base && base !== 0) ? (v$/Math.abs(base))*100 : (v$ === 0 ? 0 : null);

    const d = r[COLS.date];
    if (d) {
      const s = typeof d === 'string' ? d : (new Date(d)).toISOString();
      r.__ym = /^\d{4}-\d{2}$/.test(s.slice(0,7)) ? s.slice(0,7) : null;
    } else r.__ym = null;

    r.__isRollup = isRollupRow(r);
    r.__bucket   = bucketOf(r);
  }

  if (!all.length) throw new Error(`Alias '${alias}' returned 0 rows (after Sold/Pre-Con/Under Contract filter).`);
  return all;
}


/* =========================
   5) Aggregation
   ========================= */
function agg(rows, opts = {}) {
  // Filter out rollups to avoid double counting totals & nets
  let leaf = rows.filter(r => !r.__isRollup);
  
  // For NOI calculation (first page), include all rows (even without GL codes)
  // For all other calculations, exclude rows without GL codes
  if (!opts.noi) {
    leaf = leaf.filter(r => r[COLS.glNum] != null && String(r[COLS.glNum]).trim() !== '');
  }
  
  const byBucket = b => leaf.filter(r => r.__bucket === b);
  const sum = (arr, k) => arr.reduce((a, r) => a + (Number(r[k]) || 0), 0);

  const round2 = (x) => (x != null && Number.isFinite(x)) ? Math.round(x * 100) / 100 : x;

  if (opts.noi) {
    // First, try to find the actual "Net Operating Income (Loss)" line item
    // Search in ALL rows (including rollups) since NOI might be a calculated line item
    const noiRows = rows.filter(r => {
      const p = _canon(r[COLS.parentCat] || '');
      const c = _canon(r[COLS.cat] || '');
      // Match "Net Operating Income (Loss)" exactly or variations
      return p === 'net operating income (loss)' ||
             c === 'net operating income (loss)' ||
             (p.includes('net operating income') && p.includes('loss')) ||
             (c.includes('net operating income') && c.includes('loss'));
    });
    
    if (noiRows.length > 0) {
      // Use the actual NOI line item values (sum in case there are multiple)
      const bud = sum(noiRows, COLS.bud);
      const t12 = sum(noiRows, COLS.t12);
      const v$  = round2(t12 - bud);
      const vp  = bud ? round2((v$ / Math.abs(bud)) * 100) : (v$ === 0 ? 0 : null);
      return { bud, t12, v$, vp, bucket: 'NOI' };
    }
    
    // Fallback: calculate NOI as Income - Operating Expenses
    // For NOI, use all leaf rows (including those without GL codes) - don't filter by GL code
    const allLeaf = rows.filter(r => !r.__isRollup);
    const inc  = { 
      bud: sum(allLeaf.filter(r => r.__bucket === 'Income'), COLS.bud),  
      t12: sum(allLeaf.filter(r => r.__bucket === 'Income'), COLS.t12) 
    };
    const opex = { 
      bud: sum(allLeaf.filter(r => r.__bucket === 'Operating Expenses'), COLS.bud), 
      t12: sum(allLeaf.filter(r => r.__bucket === 'Operating Expenses'), COLS.t12) 
    };
    const bud  = inc.bud - opex.bud;
    const t12  = inc.t12 - opex.t12;
    const v$   = round2(t12 - bud);
    const vp   = bud ? round2((v$ / Math.abs(bud)) * 100) : (v$ === 0 ? 0 : null);
    return { bud, t12, v$, vp, bucket: 'NOI' };
  }

  const scoped = opts.bucket ? byBucket(opts.bucket) : leaf;
  const bud = sum(scoped, COLS.bud);
  const t12 = sum(scoped, COLS.t12);
  const v$  = round2(t12 - bud);
  const vp  = bud ? round2((v$ / Math.abs(bud)) * 100) : (v$ === 0 ? 0 : null);
  return { bud, t12, v$, vp, bucket: opts.bucket || null };
}

/* =========================
   6) Filters & grouping
   ========================= */
function currentLevel(){ return STATE.drill.length; }
function levelKey(i){ return LEVELS[i]?.key; }

function thresholds(){
  return {
    minVar: Number($('#minVar')?.value || 0),
    minPct: Number($('#minPct')?.value || 0)
  };
}

function applyFilters(baseRows){
  const ym = STATE.months?.[STATE.monthIndex] || null;
  const { minVar, minPct } = thresholds();
  const q = $('#q').value.trim().toLowerCase();

  let rows = baseRows.slice();
  // Filter by selected properties (if any are selected, otherwise show all)
  if (STATE.selectedProperties.size > 0) {
    rows = rows.filter(r => STATE.selectedProperties.has(r[COLS.prop]));
  }
  if (ym) rows = rows.filter(r => r.__ym === ym);

  for (const d of STATE.drill){
    rows = rows.filter(r => String(r[d.key] ?? '') === String(d.value));
  }

  if (q){
    rows = rows.filter(r => {
      const segs = [r[COLS.prop], r[COLS.parentCat], r[COLS.cat], r[COLS.glNum], r[COLS.glName], r[COLS.city], r[COLS.state]]
        .map(x => (x || '').toString().toLowerCase());
      return segs.some(s => s.includes(q));
    });
  }

  // Row-level threshold (kept) — helps detail screens
  rows = rows.filter(r => {
    const absVar = Math.abs(r.__var$ ?? 0);
    const absPct = Math.abs(r.__varp ?? 0);
    const okVar = (minVar <= 0) || (r.__var$ != null && absVar >= minVar);
    const okPct = (minPct <= 0) || (r.__varp == null || absPct >= minPct);
    return okVar && okPct;
  });

  return rows;
}

function groupRows(rows, level){
  const key = levelKey(level);
  if (!key) return [{ __leaf:true, rows }];

  const groups = new Map();
  for (const r of rows){
    const k = r[key] ?? '—';
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(r);
  }
  return Array.from(groups.entries()).map(([k, arr]) => ({ key, value:k, rows:arr }));
}

/* =========================
   7) Movers & helpers
   ========================= */
  function compareBarsMiniHTML(budPct, actPct){
    const clamp = v => Math.max(0, Math.min(100, v || 0));
    return `
      <div class="compare mini" title="Budget vs Actual">
        <div class="row">
          <div class="track"><div class="fill budget" style="width:${clamp(budPct)}%"></div></div>
        </div>
        <div class="row">
          <div class="track"><div class="fill actual" style="width:${clamp(actPct)}%"></div></div>
        </div>
      </div>`;
  }


   function bulletHTML(budPct, actPct, good, titleText){
  return `
    <div class="bullet ${good ? '' : 'bad'}" title="${titleText}">
      <div class="fill" style="width:${actPct}%"></div>
      <div class="target" style="left:${budPct}%"></div>
    </div>`;
}

function topCategories(rows, n=7){
  const map = new Map();
  rows.filter(r=>!r.__isRollup).forEach(r=>{
    const key = r[COLS.cat] || '—';
    const cur = map.get(key) || { name:key, v$:0, bucket:r.__bucket, parent:r[COLS.parentCat] };
    cur.v$ += (r.__var$ || 0);
    cur.bucket ||= r.__bucket;
    cur.parent ||= r[COLS.parentCat];
    map.set(key, cur);
  });
  const list = Array.from(map.values());
  list.sort((a,b)=>Math.abs(b.v$)-Math.abs(a.v$));
  return list.slice(0,n);
}
function topPropertiesByNOI(rows, n=5){
  const byProp = groupRows(rows, 0).map(g=>({ prop: g.value, ...agg(g.rows, {noi:true}) }));
  byProp.sort((a,b)=>Math.abs(b.v$)-Math.abs(a.v$));
  return byProp.slice(0,n).map(x=>x.prop);
}

/* T12 sort helpers */
function orderIndexForBucket(v){
  const i = ORDER_BUCKET.indexOf(v); return i === -1 ? 999 : i;
}
function orderIndexForParent(v){
  const i = ORDER_PARENT.indexOf(_canon(v)); return i === -1 ? 500 : i;
}

/* =========================
   8) Overview, crumbs, modals
   ========================= */
function ensureMixHeader(){
  if ($('#mixHeader')) return;
  const board = $('#board');
  if (!board) return;

  const wrap = document.createElement('div');
  wrap.id = 'mixHeader';
  wrap.style.margin = '0 0 10px 0';

  const title = document.createElement('h2');
  title.id = 'mixTitle';
  title.textContent = 'Portfolio Mix — Budget vs Actual';
  title.style.cssText = 'margin:0 0 6px 0;font-weight:800;font-size:20px;color:#243024;letter-spacing:.2px;';

  const crumbs = document.createElement('div');
  crumbs.id = 'crumbbar';
  crumbs.style.cssText = 'display:flex;gap:10px;flex-wrap:wrap;align-items:center;color:#425148;font-size:13px;';

  wrap.appendChild(title);
  wrap.appendChild(crumbs);
  board.parentElement.insertBefore(wrap, board);
}

function formatMonthLabel(ym){
  if (!ym) return '—';
  const [y, m] = ym.split('-');
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${monthNames[parseInt(m) - 1]} ${y}`;
}

function updateMonthNavButtons(){
  const prevBtn = $('#monthPrev');
  const nextBtn = $('#monthNext');
  if (prevBtn) prevBtn.disabled = STATE.monthIndex <= 0;
  if (nextBtn) nextBtn.disabled = STATE.monthIndex >= (STATE.months.length - 1);
}

function updateQuickFilterActive(){
  const minVar = Number($('#minVar')?.value || 0);
  const minPct = Number($('#minPct')?.value || 0);
  document.querySelectorAll('.quick-filter-btn').forEach(btn => {
    const btnVar = Number(btn.dataset.var || 0);
    const btnPct = Number(btn.dataset.pct || 0);
    btn.classList.toggle('active', btnVar === minVar && btnPct === minPct);
  });
}

function renderOverview(rows){
  const { bud, t12, v$, vp } = agg(rows, { noi: true });
  const properties = Array.from(new Set(rows.map(r => r[COLS.prop]).filter(Boolean)));
  const gls = Array.from(new Set(rows.map(r => r[COLS.glNum]).filter(Boolean)));

  const box = $('#summary'); if (!box) return;
  const isGood = v$ >= 0;
  box.innerHTML = `
    <div class="kpi">
      <div class="k">NOI Budget</div>
      <div class="v">${fmtMoney(bud)}</div>
    </div>
    <div class="kpi">
      <div class="k">NOI Actual</div>
      <div class="v">${fmtMoney(t12)}</div>
    </div>
    <div class="kpi ${isGood ? 'good' : 'bad'}">
      <div class="k">Variance $</div>
      <div class="v">${fmtMoney(v$)}</div>
    </div>
    <div class="kpi ${isGood ? 'good' : 'bad'}">
      <div class="k">Variance %</div>
      <div class="v">${fmtPct(vp)}</div>
    </div>
    <div class="kpi linklike" id="kpiProps">
      <div class="k">Properties</div>
      <div class="v">${properties.length}</div>
    </div>
    <div class="kpi linklike" id="kpiGLs">
      <div class="k">GL Codes</div>
      <div class="v">${gls.length}</div>
    </div>
  `;

  const movers = topCategories(rows, 7);
  $('#movers').innerHTML = movers.length
    ? movers.map(x=>{
        const good = isGoodVariance(x.bucket, x.v$, x.parent, x.name);
        return `<li>${x.name || '—'} <span class="pill ${good?'good':'bad'}" title="Good/Bad based on bucket & rules">${fmtMoney(x.v$)}</span></li>`;
      }).join('')
    : '<li>—</li>';

  $('#kpiProps')?.addEventListener('click', ()=> openPropertiesModal(rows));
  $('#kpiGLs')?.addEventListener('click',   ()=> openGLCodesModal(rows));
  
  // Render insights panel
  renderInsights(rows);
}

function renderInsights(rows){
  const panel = $('#insightsPanel'); if (!panel) return;
  
  const { bud, t12, v$, vp } = agg(rows, { noi: true });
  
  // Bucket breakdown
  const buckets = ['Income', 'Operating Expenses', 'Non-Operating', 'Debt Service', 'Other'];
  const bucketData = buckets.map(b => {
    const a = agg(rows, { bucket: b });
    return { ...a, name: b };
  }).filter(b => Math.abs(b.bud) + Math.abs(b.t12) > 0.01);
  
  // Count significant variances
  const significantVars = rows.filter(r => !r.__isRollup && Math.abs(r.__var$ || 0) >= 5000);
  const goodVarsCount = significantVars.filter(r => {
    const bucket = r.__bucket;
    return isGoodVariance(bucket, r.__var$, r[COLS.parentCat], r[COLS.cat]);
  }).length;
  const badVarsCount = significantVars.length - goodVarsCount;
  
  // Top property by variance
  const propGroups = groupRows(rows, 0);
  const propVariances = propGroups.map(g => ({
    prop: g.value,
    ...agg(g.rows, { noi: true })
  })).sort((a, b) => Math.abs(b.v$) - Math.abs(a.v$));
  const topProp = propVariances[0];
  
  const currentMonth = STATE.months[STATE.monthIndex] || '—';
  const monthLabel = formatMonthLabel(currentMonth);
  
  panel.innerHTML = `
    <div class="insight-card">
      <div class="insight-header">
        <span class="insight-icon">📊</span>
        <span class="insight-title">${monthLabel} Insights</span>
      </div>
      <div class="insight-content">
        <div class="insight-item">
          <span class="insight-label">Significant Variances:</span>
          <span class="insight-value">${significantVars.length} total (${goodVarsCount} favorable, ${badVarsCount} unfavorable)</span>
        </div>
        ${topProp ? `
        <div class="insight-item">
          <span class="insight-label">Largest Variance:</span>
          <span class="insight-value">${topProp.prop} — <span class="pill ${topProp.v$ >= 0 ? 'good' : 'bad'}">${fmtMoney(topProp.v$)}</span></span>
        </div>
        ` : ''}
        <div class="insight-item">
          <span class="insight-label">NOI Performance:</span>
          <span class="insight-value ${v$ >= 0 ? 'good' : 'bad'}">${v$ >= 0 ? 'Above' : 'Below'} budget by ${fmtMoney(Math.abs(v$))} (${fmtPct(Math.abs(vp || 0))})</span>
        </div>
      </div>
    </div>
  `;
}

function renderCrumbs(){
  ensureMixHeader();
  const host = $('#crumbbar'); if (!host) return;

  if (!STATE.drill.length){
    host.innerHTML = `<span class="crumb active">All</span>`;
    return;
  }

  const parts = [];
  parts.push(`<span class="crumb"><a href="#" data-jump="-1">All</a></span>`);
  STATE.drill.forEach((d,i)=>{
    const label = LEVELS[i]?.label || 'Level';
    const active = (i === STATE.drill.length-1) ? 'active' : '';
    parts.push(`<span class="crumb ${active}"><a href="#" data-jump="${i}">${label}: <b>${String(d.value)}</b></a></span>`);
  });
  host.innerHTML = parts.join(' ');

  host.querySelectorAll('a[data-jump]').forEach(a=>{
    a.addEventListener('click', ev=>{
      ev.preventDefault();
      const idx = Number(a.dataset.jump);
      if (idx === -1) STATE.drill = [];
      else STATE.drill = STATE.drill.slice(0, idx+1);
      render();
    });
  });
}

/* ---------- Fixed-position modal helpers ---------- */
function compareBarsHTML(budPct, actPct, bud, act){
  return `
    <div class="compare" title="Budget ${fmtMoney(bud)} • Actual ${fmtMoney(act)}">
      <div class="row">
        <div class="lbl">Budget</div>
        <div class="track"><div class="fill budget" style="width:${Math.max(0, Math.min(100, budPct))}%"></div></div>
        <div class="val">${fmtMoney(bud)}</div>
      </div>
      <div class="row">
        <div class="lbl">Actual</div>
        <div class="track"><div class="fill actual" style="width:${Math.max(0, Math.min(100, actPct))}%"></div></div>
        <div class="val">${fmtMoney(act)}</div>
      </div>
    </div>`;
}

function makeOverlay(){
  const o = document.createElement('div');
  o.style.cssText = `
    position:fixed; inset:0; background:rgba(0,0,0,.18);
    z-index:9999; display:flex; align-items:flex-start; justify-content:center;
    padding:48px 24px; overflow:auto;`;
  o.addEventListener('click', ev => { if (ev.target === o) o.remove(); });
  document.addEventListener('keydown', function esc(e){
    if (e.key === 'Escape') { o.remove(); document.removeEventListener('keydown', esc); }
  });
  return o;
}
function makeModal(title, withSearch){
  const b = document.createElement('div');
  b.style.cssText = `
    background:#fff; color:#243024; border-radius:12px; width:min(960px, calc(100% - 48px));
    max-height:calc(100% - 96px); overflow:auto; box-shadow:0 10px 30px rgba(0,0,0,.15);`;
  const head = document.createElement('div');
  head.style.cssText = 'display:flex;gap:12px;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid #e6ece8;';
  head.innerHTML = `
    <div style="font-weight:800">${title}</div>
    <div style="display:flex;gap:8px;align-items:center;">
      ${withSearch ? '<input type="text" placeholder="Search..." style="border:1px solid #dbe5df;border-radius:8px;padding:6px 10px;outline:none">' : ''}
      <button data-close class="btn btn-small" style="min-width:72px">Close</button>
    </div>`;
  const body = document.createElement('div');
  body.style.cssText = 'padding:12px 16px;';
  b.appendChild(head); b.appendChild(body);
  head.querySelector('[data-close]').addEventListener('click', ()=> b.parentElement?.remove());
  return { box:b, body, search: withSearch ? head.querySelector('input') : null };
}

/* Properties modal */
function openPropertiesModal(rows){
  const byProp = groupRows(rows, 0);
  const items = byProp
    .map(g => {
      const m = propertyMeta(g.rows);
      const u = (m.units != null && !Number.isNaN(m.units)) ? ` • ${m.units} Units` : '';
      const loc = (m.city || m.state) ? ` • ${m.city || ''}${m.city && m.state ? ', ' : ''}${m.state || ''}` : '';
      const st = m.status ? ` • ${m.status}` : '';
      return `${g.value}${u}${loc}${st}`;
    })
    .sort((a,b)=> (a>b)-(a<b));

  const overlay = makeOverlay();
  const { box, body, search } = makeModal('Properties', true);
  overlay.appendChild(box);
  document.body.appendChild(overlay);

  const ul = document.createElement('ul');
  ul.style.cssText = 'margin:0;padding-left:20px;line-height:1.75;';
  const render = (q='')=>{
    const lower = q.trim().toLowerCase();
    ul.innerHTML = items.filter(x=>!lower || x.toLowerCase().includes(lower))
      .map(x=>`<li>${x}</li>`).join('') || '<li>—</li>';
  };
  render();
  search.addEventListener('input', e=>render(e.target.value));
  body.appendChild(ul);
}

/* GL codes modal, grouped by Parent Category */
function openGLCodesModal(rows){
  const glMap = new Map();
  rows.forEach(r=>{
    if (!r[COLS.glNum]) return;
    const parent = r[COLS.parentCat] || '—';
    const label = `${r[COLS.glNum]} — ${r[COLS.glName] || ''}`;
    if (!glMap.has(parent)) glMap.set(parent, new Set());
    glMap.get(parent).add(label);
  });
  const groups = Array.from(glMap.entries())
    .map(([group, set]) => ({ group, items: Array.from(set).sort((a,b)=> (a>b)-(a<b)) }))
    .sort((a,b)=> orderIndexForParent(a.group) - orderIndexForParent(b.group) || (a.group>b.group)-(a.group<b.group));

  const overlay = makeOverlay();
  const { box, body, search } = makeModal('GL Codes by Parent Category', true);
  overlay.appendChild(box);
  document.body.appendChild(overlay);

  const draw = (q='')=>{
    const lower = q.trim().toLowerCase();
    body.innerHTML = groups.map(g=>{
      const filtered = g.items.filter(i => !lower || i.toLowerCase().includes(lower) || g.group.toLowerCase().includes(lower));
      if (!filtered.length) return '';
      return `
        <div style="margin:8px 0 4px;font-weight:700">${g.group}</div>
        <ul style="margin:0 0 12px;padding-left:20px;line-height:1.7">
          ${filtered.map(i=>`<li>${i}</li>`).join('')}
        </ul>`;
    }).join('') || '<div>—</div>';
  };
  draw();
  search.addEventListener('input', e=>draw(e.target.value));
}

/* =========================
   9) Table / cards
   ========================= */
function insertLegend(container){
  const legend = document.createElement('div');
  legend.className = 'legend';
  legend.style.cssText = 'display:flex;flex-wrap:wrap;gap:12px;margin:8px 0 4px;color:#4b5a50;font-size:12px;';
  legend.innerHTML = `
    <div style="width:180px">
      <div class="compare">
        <div class="row">
          <div class="lbl">Budget</div>
          <div class="track"><div class="fill budget" style="width:70%"></div></div>
          <div class="val">$—</div>
        </div>
        <div class="row">
          <div class="lbl">Actual</div>
          <div class="track"><div class="fill actual" style="width:55%"></div></div>
          <div class="val">$—</div>
        </div>
      </div>
    </div>
    <div>
      <span class="pill good" style="margin-right:6px">Good</span>
      <span class="pill bad">Bad</span>
    </div>`;
  container.appendChild(legend);
}


function renderTable(groups, level) {
  const container = $('#board'); if (!container) return;
  container.innerHTML = '';

  insertLegend(container);

  const { minVar, minPct } = thresholds();

  const hasData = (a) => Math.abs(a.bud || 0) + Math.abs(a.t12 || 0) > 0.00001;
  const passThresh = (a) => {
    const okVar = (minVar <= 0) || Math.abs(a.v$ || 0) >= minVar;
    const okPct = (minPct <= 0) || a.vp == null || Math.abs(a.vp || 0) >= minPct;
    return okVar && okPct;
  };

  const compact = document.documentElement.clientWidth < 1100;
  const headerName = level < LEVELS.length - 1 ? (LEVELS[level]?.label || 'Level') : 'Detail';

  const getAgg = (g) => {
    if (level === 0) return agg(g.rows, { noi: true });
    if (LEVELS[level]?.key === '__bucket') return agg(g.rows, { bucket: g.value });
    return agg(g.rows);
  };
  const nameFor = (g) => {
    if (g.__leaf) {
      const prev = LEVELS[level - 1];
      return prev?.nameFn ? prev.nameFn(g.rows[0]) : (g.rows[0]?.[prev?.key] ?? '—');
    }
    const cur = LEVELS[level];
    return cur?.nameFn ? cur.nameFn(g.rows[0]) : (g.value ?? '—');
  };
  const bucketForGroup = (g) => {
    if (level === 0) return 'NOI';
    if (LEVELS[level]?.key === '__bucket') return g.value;
    return g.rows[0]?.__bucket ?? null;
  };

  /* ---- order groups by variance magnitude (default) ---- */
  if (level === 0) {
    // Sort properties by absolute NOI variance (largest first)
    groups.sort((a,b)=>{
      const aAgg = agg(a.rows, { noi: true });
      const bAgg = agg(b.rows, { noi: true });
      return Math.abs(bAgg.v$ || 0) - Math.abs(aAgg.v$ || 0);
    });
  } else if (LEVELS[level]?.key === '__bucket') {
    groups.sort((a,b)=> orderIndexForBucket(a.value) - orderIndexForBucket(b.value) || (a.value>b.value)-(a.value<b.value));
  } else if (LEVELS[level]?.key === COLS.parentCat) {
    // Sort parent categories by absolute variance
    groups.sort((a,b)=>{
      const aAgg = getAgg(a);
      const bAgg = getAgg(b);
      const varDiff = Math.abs(bAgg.v$ || 0) - Math.abs(aAgg.v$ || 0);
      if (Math.abs(varDiff) > 0.01) return varDiff;
      return orderIndexForParent(a.value) - orderIndexForParent(b.value) || (a.value>b.value)-(a.value<b.value);
    });
  } else {
    // Sort by absolute variance for other levels
    groups.sort((a,b)=>{
      const aAgg = getAgg(a);
      const bAgg = getAgg(b);
      const varDiff = Math.abs(bAgg.v$ || 0) - Math.abs(aAgg.v$ || 0);
      if (Math.abs(varDiff) > 0.01) return varDiff;
      return (a.value>b.value)-(a.value<b.value);
    });
  }

  /* ---- drop groups that don't meet thresholds or have no data ---- */
  groups = groups.filter(g => {
    const a = getAgg(g);
    return hasData(a) && passThresh(a);
  });

  const max = Math.max(
    1,
    ...groups.map(g => {
      const a = getAgg(g);
      return Math.max(Math.abs(a.bud || 0), Math.abs(a.t12 || 0));
    })
  );

  const topOffProps = (level === 0) ? new Set(topPropertiesByNOI( applyFilters(STATE.rows), 5 )) : new Set();

  if (compact) {
    groups.forEach(g => {
      const a = getAgg(g);
      const name = nameFor(g);
      const budPct = Math.min(100, Math.abs(a.bud || 0) / max * 100);
      const actPct = Math.min(100, Math.abs(a.t12 || 0) / max * 100);
      const bucket = bucketForGroup(g);
      const rep = g.rows[0] || {};
      const good = isGoodVariance(bucket, a.v$, rep[COLS.parentCat], rep[COLS.cat]);

      const meta = (level === 0) ? propertyMeta(g.rows) : null;
      const metaLine = meta ? `
        <div class="sub">
          ${meta.units != null ? `${meta.units} Units` : ''}${meta.units && (meta.city||meta.state) ? ' • ' : ''}
          ${(meta.city || '')}${meta.city && meta.state ? ', ' : ''}${meta.state || ''}${meta.status ? ` • ${meta.status}` : ''}
        </div>` : '';

     const bullet = bulletHTML(
        budPct, actPct, good,
        `Budget ${fmtMoney(a.bud)} • Actual ${fmtMoney(a.t12)}`
      );

     const viz = compareBarsHTML(budPct, actPct, a.bud, a.t12);

      const card = document.createElement('div');
      card.className = 'rowcard';
      const highVarianceBadge = (level === 0 && topOffProps.has(name)) 
        ? `<span class="pill ${good ? 'good' : 'bad'}" style="margin-left:6px">High Variance</span>`
        : '';
      card.innerHTML = `
        <div>
          <div class="title">${name || '—'} ${highVarianceBadge}</div>
          ${metaLine}
        </div>

        <div class="kvs">
          <div class="kv"><div class="k">${level===0?'NOI Actual':'Actual'}</div><div class="v">${fmtMoney(a.t12)}</div></div>
          <div class="kv"><div class="k">Variance $</div><div class="v"><span class="pill ${good?'good':'bad'}">${fmtMoney(a.v$)}</span></div></div>
          <div class="kv"><div class="k">${level===0?'NOI Budget':'Budget'}</div><div class="v">${fmtMoney(a.bud)}</div></div>
          <div class="kv"><div class="k">Variance %</div><div class="v"><span class="pill ${good?'good':'bad'}">${fmtPct(a.vp)}</span></div></div>
        </div>

        ${viz}
      `;


      if (level < LEVELS.length - 1) {
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => {
          STATE.drill.push({ key: LEVELS[level].key, value: g.value });
          render();
        });
      }
      container.appendChild(card);
    });
    return;
  }

  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr>
        <th>${headerName}</th>
        <th>${level===0 ? 'NOI: Budget vs Actual' : 'Budget vs Actual'}</th>
        <th class="num sort" data-k="bud">${level===0 ? 'NOI Budget' : 'Budget'}</th>
        <th class="num sort" data-k="t12">${level===0 ? 'NOI Actual' : 'Actual'}</th>
        <th class="num sort" data-k="v$">Variance $</th>
        <th class="num sort" data-k="vp">Variance %</th>
        <th class="num">Rows</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  const tb = table.querySelector('tbody');

  const drawRow = (g) => {
    const a = getAgg(g);
    const name = nameFor(g);
    const budPct = Math.min(100, Math.abs(a.bud || 0) / max * 100);
    const actPct = Math.min(100, Math.abs(a.t12 || 0) / max * 100);
    const bucket = bucketForGroup(g);
    const rep = g.rows[0] || {};
    const good = isGoodVariance(bucket, a.v$, rep[COLS.parentCat], rep[COLS.cat]);

    const meta = (level === 0) ? propertyMeta(g.rows) : null;
    const metaLine = meta ? `
      <div class="sub">
        ${meta.units != null ? `${meta.units} Units` : ''}${meta.units && (meta.city||meta.state) ? ' • ' : ''}
        ${(meta.city || '')}${meta.city && meta.state ? ', ' : ''}${meta.state || ''}${meta.status ? ` • ${meta.status}` : ''}
      </div>` : '';

    const tr = document.createElement('tr');
    tr.className = 'row';
    const highVarianceBadge = (level === 0 && topOffProps.has(name))
      ? `<span class="pill ${good ? 'good' : 'bad'}" style="margin-left:6px">High Variance</span>`
      : '';
    tr.innerHTML = `
      <td>
        <div>${name || '—'} ${highVarianceBadge}</div>
        ${metaLine}
      </td>
      <td class="barcell">
        ${compareBarsMiniHTML(budPct, actPct, a.bud, a.t12)}
      </td>


      <td class="num">${fmtMoney(a.bud)}</td>
      <td class="num">${fmtMoney(a.t12)}</td>
      <td class="num"><span class="pill ${good?'good':'bad'}" title="Good/Bad based on bucket & rules">${fmtMoney(a.v$)}</span></td>
      <td class="num"><span class="pill ${good?'good':'bad'}">${fmtPct(a.vp)}</span></td>
      <td class="num">${g.rows.length}</td>
    `;

    if (level < LEVELS.length - 1) {
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', () => {
        STATE.drill.push({ key: LEVELS[level].key, value: g.value });
        render();
      });
    }
    return tr;
  };

  groups.forEach(g => tb.appendChild(drawRow(g)));
  container.appendChild(table);

  // Set default sort to variance $ (descending)
  const defaultSortTh = table.querySelector('th.sort[data-k="v$"]');
  if (defaultSortTh && !defaultSortTh.dataset.dir) {
    defaultSortTh.dataset.dir = 'desc';
  }

  table.querySelectorAll('th.sort').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.k;
      const dir = th.dataset.dir === 'asc' ? 'desc' : 'asc';
      
      // Clear other sort indicators
      table.querySelectorAll('th.sort').forEach(t => {
        if (t !== th) delete t.dataset.dir;
      });
      th.dataset.dir = dir;

      groups.sort((A, B) => {
        const a = (getAgg(A)[k] || 0);
        const b = (getAgg(B)[k] || 0);
        // For variance, sort by absolute value if not already sorted
        if (k === 'v$' && dir === 'desc') {
          return Math.abs(b) - Math.abs(a);
        }
        return (dir === 'asc' ? 1 : -1) * ((a > b) - (a < b));
      });

      const newBody = document.createElement('tbody');
      groups.forEach(g => newBody.appendChild(drawRow(g)));
      table.replaceChild(newBody, tb);
    });
  });
  
  // Apply default sort on initial render
  if (defaultSortTh && !defaultSortTh.dataset.dir) {
    defaultSortTh.click();
  }
}

/* =========================
   10) Main render
   ========================= */
function render(){
  const errEl = $('#error'); if (errEl) errEl.hidden = true;

  const rows = applyFilters(STATE.rows);
  renderOverview(rows);
  renderCrumbs();

  const lvl = currentLevel();
  let groups = groupRows(rows, lvl);

  // Apply same sorting logic as renderTable
  if (lvl === 0) {
    groups.sort((a,b)=>{
      const aAgg = agg(a.rows, { noi: true });
      const bAgg = agg(b.rows, { noi: true });
      return Math.abs(bAgg.v$ || 0) - Math.abs(aAgg.v$ || 0);
    });
  } else if (LEVELS[lvl]?.key === '__bucket') {
    groups.sort((a,b)=> orderIndexForBucket(a.value) - orderIndexForBucket(b.value) || (a.value>b.value)-(a.value<b.value));
  } else if (LEVELS[lvl]?.key === COLS.parentCat) {
    groups.sort((a,b)=>{
      const aAgg = agg(a.rows);
      const bAgg = agg(b.rows);
      const varDiff = Math.abs(bAgg.v$ || 0) - Math.abs(aAgg.v$ || 0);
      if (Math.abs(varDiff) > 0.01) return varDiff;
      return orderIndexForParent(a.value) - orderIndexForParent(b.value) || (a.value>b.value)-(a.value<b.value);
    });
  } else {
    groups.sort((a,b)=>{
      const aAgg = agg(a.rows);
      const bAgg = agg(b.rows);
      const varDiff = Math.abs(bAgg.v$ || 0) - Math.abs(aAgg.v$ || 0);
      if (Math.abs(varDiff) > 0.01) return varDiff;
      return (a.value>b.value)-(a.value<b.value);
    });
  }

  renderTable(groups, lvl);

  const f = $('#footer');
  if (f) f.innerHTML = `Level: <b>${LEVELS[lvl]?.label || 'Detail'}</b> · Showing <b>${rows.length.toLocaleString()}</b> rows`;
}

/* =========================
   11) UI init
   ========================= */
function makeSliderBinding(numberInput, opts){
  // Keep the original id so the rest of the app reads #minVar / #minPct
  const id = numberInput.id || '';
  const start = numberInput.value || '0';

  const wrap = document.createElement('div');
  wrap.className = 'range-wrap';

  const range = document.createElement('input');
  range.type = 'range';
  range.min  = String(opts.min);
  range.max  = String(opts.max);
  range.step = String(opts.step);
  range.value = start;

  const numeric = document.createElement('input');
  numeric.type = 'number';
  numeric.min  = String(opts.min);
  numeric.max  = String(opts.max);
  numeric.step = String(opts.step);
  numeric.value = start;

  // One persistent hidden input that keeps the id (#minVar / #minPct)
  const hidden = document.createElement('input');
  hidden.type = 'hidden';
  if (id) hidden.id = id;
  hidden.value = start;

  // Swap the old number input for our wrapper
  numberInput.replaceWith(wrap);
  wrap.appendChild(range);
  wrap.appendChild(numeric);
  wrap.appendChild(hidden);

  // Sync all three elements and fire events so existing listeners run
  const sync = (val, fire=true) => {
    range.value   = val;
    numeric.value = val;
    hidden.value  = val;

    if (fire) {
      // keep the existing event wiring working
      hidden.dispatchEvent(new Event('input',  { bubbles: true }));
      hidden.dispatchEvent(new Event('change', { bubbles: true }));
      // and force an immediate refresh for good measure
      if (typeof render === 'function') render();
    }
  };

  range.addEventListener('input',   e => sync(e.target.value));
  numeric.addEventListener('input', e => sync(e.target.value));

  // Initialize without double-firing
  sync(start, false);
}

/* =========================
   Property Multi-Select
   ========================= */
function initPropertyMultiSelect(rows){
  // Get all unique properties sorted by BirthOrder (rows should already be filtered by current month)
  const propOrder = new Map();
  for (const r of rows) {
    const p = r[COLS.prop]; if (!p) continue;
    const b = Number.isFinite(r[COLS.birth]) ? r[COLS.birth] : Infinity;
    const prev = propOrder.get(p);
    if (prev == null || b < prev) propOrder.set(p, b);
  }
  let allProps = Array.from(new Set(rows.map(r => r[COLS.prop]).filter(Boolean)))
    .sort((a,b)=>{
      const A = propOrder.get(a) ?? Infinity;
      const B = propOrder.get(b) ?? Infinity;
      if (A !== B) return A - B;
      return (a > b) - (a < b);
    });

  // Only keep selected properties that exist in this month's list; default to all selected if none
  STATE.selectedProperties.forEach(p => { if (!allProps.includes(p)) STATE.selectedProperties.delete(p); });
  if (STATE.selectedProperties.size === 0) {
    allProps.forEach(p => STATE.selectedProperties.add(p));
  }

  const trigger = $('#propTrigger');
  const dropdown = $('#propDropdown');
  const display = $('#propDisplay');
  const list = $('#propList');
  const searchInput = $('#propSearch');
  const selectAllBtn = $('#propSelectAll');
  const clearAllBtn = $('#propClearAll');

  if (!trigger || !dropdown || !display || !list) return;

  // Update display text
  function updateDisplay(){
    const count = STATE.selectedProperties.size;
    const total = allProps.length;
    if (count === 0) {
      display.textContent = 'No Properties Selected';
    } else if (count === total) {
      display.textContent = 'All Properties';
    } else {
      display.textContent = `${count} of ${total} Selected`;
    }
  }

  // Render property list
  function renderPropertyList(filterText = ''){
    const filter = filterText.toLowerCase().trim();
    const filtered = filter 
      ? allProps.filter(p => p.toLowerCase().includes(filter))
      : allProps;

    if (filtered.length === 0) {
      list.innerHTML = '<div class="property-multiselect-empty">No properties found</div>';
      return;
    }

    list.innerHTML = filtered.map(prop => {
      const checked = STATE.selectedProperties.has(prop);
      const safeId = prop.replace(/[^a-zA-Z0-9]/g, '-');
      return `
        <div class="property-multiselect-item ${checked ? 'checked' : ''}">
          <input type="checkbox" id="prop-${safeId}" 
                 data-prop="${prop}" ${checked ? 'checked' : ''}>
          <label for="prop-${safeId}">${prop}</label>
        </div>
      `;
    }).join('');

    // Add event listeners to checkboxes
    list.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', (e) => {
        const prop = e.target.dataset.prop;
        const item = e.target.closest('.property-multiselect-item');
        if (e.target.checked) {
          STATE.selectedProperties.add(prop);
          item?.classList.add('checked');
        } else {
          STATE.selectedProperties.delete(prop);
          item?.classList.remove('checked');
        }
        updateDisplay();
        render();
      });
    });
  }

  // Toggle dropdown (use flex when open so it overlays; none when closed so it doesn't affect layout)
  function openDropdown() {
    dropdown.style.display = 'flex';
    dropdown.style.visibility = 'visible';
    trigger.classList.add('active');
    searchInput.value = '';
    renderPropertyList();
    searchInput.focus();
  }
  function closeDropdown() {
    dropdown.style.display = 'none';
    dropdown.style.visibility = '';
    trigger.classList.remove('active');
  }
  trigger.addEventListener('click', (e) => {
    e.stopPropagation();
    const isOpen = dropdown.style.display === 'flex';
    if (isOpen) closeDropdown();
    else openDropdown();
  });

  // Close dropdown when clicking outside or pressing Escape
  document.addEventListener('click', (e) => {
    if (!dropdown.contains(e.target) && e.target !== trigger) {
      closeDropdown();
    }
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && dropdown.style.display === 'flex') {
      closeDropdown();
      trigger.focus();
    }
  });

  // Search functionality
  searchInput.addEventListener('input', (e) => {
    renderPropertyList(e.target.value);
  });

  // Select All
  selectAllBtn.addEventListener('click', () => {
    const filter = searchInput.value.toLowerCase().trim();
    const filtered = filter 
      ? allProps.filter(p => p.toLowerCase().includes(filter))
      : allProps;
    filtered.forEach(p => STATE.selectedProperties.add(p));
    renderPropertyList(searchInput.value);
    updateDisplay();
    render();
  });

  // Clear All
  clearAllBtn.addEventListener('click', () => {
    const filter = searchInput.value.toLowerCase().trim();
    const filtered = filter 
      ? allProps.filter(p => p.toLowerCase().includes(filter))
      : allProps;
    filtered.forEach(p => STATE.selectedProperties.delete(p));
    renderPropertyList(searchInput.value);
    updateDisplay();
    render();
  });

  // Refresh property list when month changes: only show properties that have data for current month
  _refreshPropertyListForMonth = function(){
    const ym = STATE.months?.[STATE.monthIndex] || null;
    const rowsForMonth = ym && STATE.rows.length ? STATE.rows.filter(r => r.__ym === ym) : STATE.rows;
    const propOrder = new Map();
    for (const r of rowsForMonth) {
      const p = r[COLS.prop]; if (!p) continue;
      const b = Number.isFinite(r[COLS.birth]) ? r[COLS.birth] : Infinity;
      const prev = propOrder.get(p);
      if (prev == null || b < prev) propOrder.set(p, b);
    }
    allProps = Array.from(new Set(rowsForMonth.map(r => r[COLS.prop]).filter(Boolean)))
      .sort((a,b)=>{
        const A = propOrder.get(a) ?? Infinity;
        const B = propOrder.get(b) ?? Infinity;
        if (A !== B) return A - B;
        return (a > b) - (a < b);
      });
    STATE.selectedProperties.forEach(p => { if (!allProps.includes(p)) STATE.selectedProperties.delete(p); });
    if (STATE.selectedProperties.size === 0) allProps.forEach(p => STATE.selectedProperties.add(p));
    updateDisplay();
    renderPropertyList();
  };

  // Initial render
  updateDisplay();
  renderPropertyList();
}

function initUI(rows){
  const lu = $('#lastUpdated'); if (lu) lu.textContent = `Last updated ${new Date().toLocaleString()}`;

  // Hide/remove the old Variance Direction block (if still in HTML)
  const dirBlock = Array.from(document.querySelectorAll('.ctrl'))
    .find(el => /variance direction/i.test(el.textContent || ''));
  if (dirBlock) dirBlock.style.display = 'none';

  // Slider bindings
  const minVarEl = document.getElementById('minVar'); if (minVarEl) makeSliderBinding(minVarEl, { min:0, max:200000, step:100 });
  const minPctEl = document.getElementById('minPct'); if (minPctEl) makeSliderBinding(minPctEl, { min:0, max:100, step:0.1 });

  // Months (set first so property list can be filtered by current month)
  const months = Array.from(new Set(rows.map(r => r.__ym).filter(Boolean))).sort();
  STATE.months = months;
  const latestIdx = months.length ? months.length - 1 : 0;
  STATE.monthIndex = latestIdx;
  $('#month').innerHTML = (months.length
    ? months.map((m,i)=>`<option value="${i}" ${i===latestIdx?'selected':''}>${formatMonthLabel(m)}</option>`).join('')
    : `<option value="">(no months)</option>`);

  // Property multi-select: only properties that have data for the selected month
  const ym = STATE.months[STATE.monthIndex] || null;
  const rowsForMonth = ym ? rows.filter(r => r.__ym === ym) : rows;
  initPropertyMultiSelect(rowsForMonth);

  // Month navigation
  const monthPrev = $('#monthPrev');
  const monthNext = $('#monthNext');
  const refreshPropsThenRender = () => { _refreshPropertyListForMonth && _refreshPropertyListForMonth(); render(); };
  if (monthPrev) {
    monthPrev.addEventListener('click', () => {
      if (STATE.monthIndex > 0) {
        STATE.monthIndex--;
        $('#month').value = STATE.monthIndex;
        updateMonthNavButtons();
        refreshPropsThenRender();
      }
    });
  }
  if (monthNext) {
    monthNext.addEventListener('click', () => {
      if (STATE.monthIndex < months.length - 1) {
        STATE.monthIndex++;
        $('#month').value = STATE.monthIndex;
        updateMonthNavButtons();
        refreshPropsThenRender();
      }
    });
  }
  updateMonthNavButtons();

  // Quick filter buttons
  document.querySelectorAll('.quick-filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const varVal = btn.dataset.var || '0';
      const pctVal = btn.dataset.pct || '0';
      $('#minVar').value = varVal;
      $('#minPct').value = pctVal;
      // Trigger sync for sliders
      $('#minVar').dispatchEvent(new Event('input', { bubbles: true }));
      $('#minPct').dispatchEvent(new Event('input', { bubbles: true }));
      // Update active state
      document.querySelectorAll('.quick-filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      render();
    });
  });

  // Events
  const evts = ['input','change'];
  // Property multi-select events are handled in initPropertyMultiSelect
  evts.forEach(e => $('#month').addEventListener(e, ev => { 
    STATE.monthIndex = Number(ev.target.value||0); 
    updateMonthNavButtons();
    _refreshPropertyListForMonth && _refreshPropertyListForMonth();
    render(); 
  }));
  evts.forEach(e => $('#minVar').addEventListener(e, () => {
    updateQuickFilterActive();
    render();
  }));
  evts.forEach(e => $('#minPct').addEventListener(e, () => {
    updateQuickFilterActive();
    render();
  }));
  evts.forEach(e => $('#q').addEventListener(e, render));
  $('#exportBtn')?.addEventListener('click', exportCSV);

  const setCompact = () => document.documentElement.classList.toggle('compact', window.innerWidth < 1100);
  setCompact(); window.addEventListener('resize', setCompact);

  ensureMixHeader();
  updateQuickFilterActive();
  render();
}

/* =========================
   12) Export
   ========================= */
function exportCSV(){
  const rows = applyFilters(STATE.rows);
  const cols = [COLS.prop, COLS.date, COLS.parentCat, COLS.cat, COLS.glNum, COLS.glName, COLS.bud, COLS.t12];
  const header = cols.concat(['variance$', 'variance%']);
  const lines = [header.join(',')];
  for (const r of rows){
    const cells = [
      r[COLS.prop]||'',
      r[COLS.date]||'',
      r[COLS.parentCat]||'',
      r[COLS.cat]||'',
      r[COLS.glNum]||'',
      (r[COLS.glName]||'').toString().replace(/,/g,''),
      r[COLS.bud] ?? '',
      r[COLS.t12] ?? '',
      r.__var$ ?? '',
      r.__varp ?? ''
    ];
    lines.push(cells.join(','));
  }
  const blob = new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'variance_export.csv'; a.click();
  URL.revokeObjectURL(url);
}

/* =========================
   13) Boot
   ========================= */
(async function main(){
  try {
    let rows, lastErr;
    for (const alias of DATASET_ALIASES) {
      try {
        rows = await fetchRowsByAlias(alias);
        break;
      } catch (e) {
        lastErr = e;
      }
    }
    if (!rows) throw lastErr;
    const errEl = document.getElementById('error'); if (errEl) errEl.hidden = true;
    STATE.rows = rows;
    initUI(rows);
  } catch (err) {
    showError(err);
    const board = document.getElementById('board'); if (board) board.innerHTML = '';
  }
})();
