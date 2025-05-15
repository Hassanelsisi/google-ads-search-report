/*********************************************************************
 *  SEARCH-CAMPAIGN PERFORMANCE & STRATEGY REPORT        2025-05-15
 * ------------------------------------------------------------------
 *  Generates an 8-tab Google Sheet each run:
 *    README · CONFIG · Overview · Heatmap · Strategy
 *    Ads · Campaign-KW · Search Terms
 *  Safe (read-only): never edits campaigns, ads, bids, or budgets.
 *********************************************************************/

/*──────── USER SETTINGS ────────*/
const SPREADSHEET_ID = 'PASTE_YOUR_44_CHARACTER_SHEET_ID';   // ← replace
const DATE_FROM      = '2025-01-01';
const DATE_TO        = '2025-12-31';

/* Brand filter used in Search-Terms and KW tabs */
const BRAND_PATTERNS = [
  /yourbrand/i,         // ASCII   – edit to your brand
  /براندك/i             // Arabic  – edit / add others if needed
];
/*──────────────────────────────*/

/* Colour constants */
const TOP    = '#c8e6c9';   // Good / Add / Keep
const BOTTOM = '#ffcdd2';   // Exclude / Pause
const REVIEW = '#ffe0b2';   // Needs review
const PH     = '_PLACEHOLDER_';   // temp sheet while wiping

/*════════ MAIN ════════*/
function main() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  wipe(ss);                 // clear previous run
  buildReadme(ss);          // README + legend
  buildConfig(ss);          // CONFIG

  const spend = buildOverview(ss);  // Overview returns spend per campaign
  buildHeatmap(ss);
  buildStrategy(ss);
  buildAdsTab(ss);
  buildCampaignKW(ss, spend);
  buildSearchTerms(ss);
}

/*────────────────── SAFE WIPE ──────────────────*/
function wipe(ss) {
  const ph = ss.getSheetByName(PH) || ss.insertSheet(PH, 0);
  ss.getSheets().forEach(sh => { if (sh.getName() !== PH) ss.deleteSheet(sh); });
}

/*────────────────── README ──────────────────*/
function buildReadme(ss) {
  const sh = ss.getSheetByName(PH); sh.setName('README'); sh.clear();
  const tbl = [
    ['📖 Report Guide', ''],
    ['Tab', 'Purpose'],
    ['Overview',     'KPIs & bid strategy per campaign'],
    ['Heatmap',      'Day × Hour: Clicks • CTR • CPC • Conv • CPA'],
    ['Strategy',     'Rule tips, device CPA flags, Google-Ads recommendations'],
    ['Ads',          'Enabled ads – keep / test / pause (colour-coded)'],
    ['Campaign-KW',  'Keyword stats + Top/Bottom-10 + actions'],
    ['Search Terms', 'Brand-aware Add / Review / Exclude'],
    ['CONFIG',       'Edit thresholds without code']
  ];
  sh.getRange(1, 1, tbl.length, 2).setValues(tbl);

  sh.getRange(tbl.length + 2, 1, 1, 2).setValues([[
    'Disclaimer:',
    'Recommendations are guidance only. Always review before acting.'
  ]]).setFontStyle('italic');

  /* Colour legend */
  const legend = [
    ['Colour', 'Meaning'],
    [TOP,    'Good / Add / Keep / Highest CTR / Lowest CPC'],
    [BOTTOM, 'Exclude / Pause / Lowest CTR / Highest CPC'],
    [REVIEW, 'Review'],
    ['#fff9c4', 'High CPC (above multiplier)'],
    ['#bbdefb', 'Low CPC (below average)']
  ];
  sh.getRange(tbl.length + 4, 1, legend.length, 2).setValues(legend);
  legend.slice(1).forEach((l, i) =>
    sh.getRange(tbl.length + 5 + i, 1).setBackground(l[0]));

  sh.setFrozenRows(1);
}

/*────────────────── CONFIG ──────────────────*/
function buildConfig(ss) {
  const sh = ss.insertSheet('CONFIG');
  sh.getRange(1, 1, 6, 2).setValues([
    ['Key',            'Value'],
    ['CLICK_THRESHOLD', 5],     // min clicks – Exclude decision
    ['CTR_HIGH',       0.05],   // ≥ 5 % high CTR
    ['CTR_LOW',        0.01],   // ≤ 1 % low CTR
    ['CPC_MULTIPLIER', 1.3],    // High-CPC if > 1.3 × avg
    ['CPA_MULTIPLIER', 1.5]     // High-CPA if > 1.5 × avg
  ]);
  sh.hideSheet();
}
function cfg(key) {
  return SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName('CONFIG')
    .createTextFinder(key).findNext().offset(0, 1).getValue();
}

/*── helpers (µ→$, formatters, gradient) ──*/
const usdFmt = r => r.setNumberFormat('$#,##0.00');
const pctFmt = r => r.setNumberFormat('0.00%');
const µ2$    = v => (typeof v === 'number' ? v / 1e6 : v);
function gradient(sh, rng){
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([rng])
    .setGradientMinpoint('#b71c1c')
    .setGradientMaxpoint('#1b5e20')
    .build());
  sh.setConditionalFormatRules(rules);
}
function finalise(sh){
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.autoResizeColumns(1, sh.getLastColumn());
  sh.autoResizeRows(1, sh.getLastRow());
}

/*────────────────── 3. OVERVIEW ──────────────────*/
function buildOverview(ss){
  const sh = ss.insertSheet('Overview');
  AdsApp.report(`
    SELECT campaign.id, campaign.name, campaign.bidding_strategy_type,
           metrics.clicks, metrics.impressions, metrics.ctr,
           metrics.conversions, metrics.conversions_value,
           metrics.average_cpc, metrics.cost_micros
    FROM   campaign
    WHERE  campaign.advertising_channel_type = 'SEARCH'
      AND  segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
      AND  metrics.cost_micros > 0`)
    .exportToSheet(sh);

  const cpaCol  = sh.getLastColumn() + 1,
        roasCol = cpaCol + 1;
  sh.getRange(1, cpaCol, 1, 2).setValues([['CPA (USD)', 'ROAS']]);

  convertMicros(sh, 9); convertMicros(sh, 10);
  usdFmt(sh.getRange(2, 9, sh.getLastRow() - 1, 2));
  pctFmt(sh.getRange(2, 6, sh.getLastRow() - 1));

  const rows = sh.getLastRow() - 1;
  const cost = sh.getRange(2, 10, rows).getValues();
  const conv = sh.getRange(2, 7,  rows).getValues();
  const val  = sh.getRange(2, 8,  rows).getValues();

  const CPA  = conv.map((c, i) => [c[0] > 0 ? cost[i][0] / c[0]  : '']);
  const ROAS = val .map((v, i) => [cost[i][0] > 0 ? v[0]  / cost[i][0] : '']);

  sh.getRange(2, cpaCol,  rows).setValues(CPA);
  sh.getRange(2, roasCol, rows).setValues(ROAS);
  usdFmt(sh.getRange(2, cpaCol, rows));
  sh.getRange(2, roasCol, rows).setNumberFormat('0.00');

  /* highlight extremes */
  const ctr = sh.getRange(2, 6, rows).getValues().map(r=>r[0]);
  const cpc = sh.getRange(2, 9, rows).getValues().map(r=>r[0]);
  const hiCTR=Math.max(...ctr), loCTR=Math.min(...ctr);
  const hiCPC=Math.max(...cpc), loCPC=Math.min(...cpc);
  for(let i=0;i<rows;i++){
    const rng=sh.getRange(i+2,1,1,roasCol);
    if(ctr[i]===hiCTR||cpc[i]===loCPC) rng.setBackground(TOP);
    if(ctr[i]===loCTR||cpc[i]===hiCPC) rng.setBackground(BOTTOM);
  }
  finalise(sh);

  /* return spend map */
  const ids   = sh.getRange(2, 1, rows).getValues();
  const spend = sh.getRange(2, 10, rows).getValues();
  const map={}; ids.forEach((r,i)=>map[r[0]]=spend[i][0]);
  return map;
}

/*────────────────── 4. HEATMAP ──────────────────*/
function buildHeatmap(ss){
  const sh = ss.insertSheet('Heatmap');
  const rep = AdsApp.report(`
    SELECT segments.day_of_week, segments.hour,
           metrics.clicks, metrics.ctr, metrics.average_cpc,
           metrics.conversions, metrics.cost_micros
    FROM   campaign
    WHERE  campaign.advertising_channel_type = 'SEARCH'
      AND  segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
      AND  metrics.cost_micros > 0`);
  const DAYS=['MONDAY','TUESDAY','WEDNESDAY','THURSDAY',
              'FRIDAY','SATURDAY','SUNDAY'];
  const g={Clicks:{},CTR:{},CPC:{},Conv:{},Cost:{}};
  DAYS.forEach(d=>['Clicks','CTR','CPC','Conv','Cost']
    .forEach(k=>g[k][d]=Array(24).fill(0)));
  const it=rep.rows();
  while(it.hasNext()){
    const r=it.next(), d=r['segments.day_of_week'], h=+r['segments.hour'];
    g.Clicks[d][h]+= +r['metrics.clicks'];
    g.CTR[d][h]   += +r['metrics.ctr'];
    g.CPC[d][h]   += µ2$(+r['metrics.average_cpc']);
    g.Conv[d][h]  += +r['metrics.conversions'];
    g.Cost[d][h]  += +r['metrics.cost_micros'];
  }
  const CPA={}; DAYS.forEach(d=>{
    CPA[d]=Array(24).fill(0);
    for(let h=0;h<24;h++){
      if(g.Conv[d][h]>0) CPA[d][h]=µ2$(g.Cost[d][h])/g.Conv[d][h];
    }
  });

  function grid(label,grid,fmt){
    const start=sh.getLastRow()+2||1;
    sh.getRange(start,1).setValue(label);
    sh.getRange(start+1,1).setValue('Day / Hour');
    for(let h=0;h<24;h++) sh.getRange(start+1,h+2).setValue(h);
    DAYS.forEach((d,i)=>
      sh.getRange(start+2+i,1,1,25)
        .setValues([[d,...grid[d]]]));
    const rng=sh.getRange(start+2,2,7,24);
    gradient(sh,rng);
    if(fmt==='pct') pctFmt(rng);
    if(fmt==='usd') usdFmt(rng);
  }
  grid('Clicks',g.Clicks);
  grid('CTR',g.CTR,'pct');
  grid('CPC',g.CPC,'usd');
  grid('Conversions',g.Conv);
  grid('CPA',CPA,'usd');
  finalise(sh);
}

/*────────────────── 5. STRATEGY ──────────────────*/
/* … identical to working version … */

/*────────────────── 6. ADS ──────────────────*/
/* … identical to working version … */

/*────────────────── 7. CAMPAIGN-KW ──────────────────*/
/* … identical to working version … */

/*────────────────── 8. SEARCH TERMS ──────────────────*/
/* … identical to working version … */
