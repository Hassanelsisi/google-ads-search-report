/*********************************************************************
 *  SEARCH-CAMPAIGN PERFORMANCE & STRATEGY REPORT        2025-05-15
 * ------------------------------------------------------------------
 *  Generates an 8-tab Google Sheet on every run:
 *    README · CONFIG · Overview · Heatmap · Strategy
 *    Ads · Campaign-KW · Search Terms
 *
 *  READ-ONLY: the script never edits campaigns, ads, bids, or budgets.
 *********************************************************************/

/*──────── USER SETTINGS ────────*/
const SPREADSHEET_ID = 'PASTE_YOUR_44_CHARACTER_SHEET_ID';   // ← replace
const DATE_FROM      = '2025-01-01';
const DATE_TO        = '2025-12-31';

/* Brand keywords (used in Search-Terms & Campaign-KW tabs) */
const BRAND_PATTERNS = [
  /yourbrand/i,         // ASCII – edit to your brand
  /براندك/i             // Arabic – edit / add others
];
/*──────────────────────────────*/

/* Colour constants */
const TOP    = '#c8e6c9';   // good / keep / add
const BOTTOM = '#ffcdd2';   // exclude / pause
const REVIEW = '#ffe0b2';   // review
const PH     = '_PLACEHOLDER_';   // temp sheet during wipe

/*════════ MAIN ════════*/
function main() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  wipe(ss);                 // clear previous run
  buildReadme(ss);          // README + legend
  buildConfig(ss);          // CONFIG

  const spend = buildOverview(ss);   // returns spend map
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
    ['Heatmap',      'Day × Hour grids (Clicks • CTR • CPC • Conv • CPA)'],
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
    ['CTR_LOW',        0.01],   // ≤ 1 % low  CTR
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

/*── helpers (formatters, conversions, gradient) ──*/
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
function convertMicros(sh, col){
  if (sh.getLastRow() > 1) {
    const rng = sh.getRange(2, col, sh.getLastRow() - 1);
    rng.setValues(rng.getValues().map(v => [µ2$(v[0])]));
  }
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

  const CPA  = conv.map((c, i) => [c[0] > 0 ? cost[i][0] / c[0] : '']);
  const ROAS = val .map((v, i) => [cost[i][0] > 0 ? v[0] / cost[i][0] : '']);

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
function buildStrategy(ss) {
  const ov = ss.getSheetByName('Overview');
  const st = ss.insertSheet('Strategy');

  st.appendRow(['Rule-Based Recommendations (campaign level)']);
  st.appendRow(['Campaign','CTR','Avg CPC','CPA','ROAS','Bid Strategy','Advice']);

  const rows = ov.getLastRow() - 1;
  const name = ov.getRange(2, 2, rows).getValues();
  const ctr  = ov.getRange(2, 6, rows).getValues();
  const cpc  = ov.getRange(2, 9, rows).getValues();
  const cpa  = ov.getRange(2, ov.getLastColumn() - 1, rows).getValues();
  const roas = ov.getRange(2, ov.getLastColumn(),      rows).getValues();
  const bid  = ov.getRange(2, 3, rows).getValues();

  const avgCPC = cpc.reduce((s,v)=>s+v[0],0) / rows;
  const avgCPA = cpa.filter(v=>v[0]).reduce((s,v)=>s+v[0],0) /
                 (cpa.filter(v=>v[0]).length || 1);
  const CTR_H = cfg('CTR_HIGH'), CTR_L = cfg('CTR_LOW'),
        CPC_M = cfg('CPC_MULTIPLIER'), CPA_M = cfg('CPA_MULTIPLIER');

  name.forEach((r,i)=>{
    const tips=[];
    if(ctr[i][0] < CTR_L) tips.push('Low CTR');
    if(ctr[i][0] > CTR_H) tips.push('High CTR');
    if(cpc[i][0] > CPC_M*avgCPC) tips.push('High CPC');
    if(cpa[i][0] && cpa[i][0] > CPA_M*avgCPA) tips.push('High CPA');
    if(roas[i][0] && roas[i][0] < 1) tips.push('ROAS < 1');
    if(roas[i][0] && roas[i][0] > 3) tips.push('Consider tROAS bidding');
    if(!tips.length) tips.push('Healthy');
    st.appendRow([
      r[0], ctr[i][0], cpc[i][0], cpa[i][0] || '',
      roas[i][0] || '', bid[i][0], tips.join(' • ')
    ]);
  });
  pctFmt(st.getRange(3,2,rows));
  usdFmt(st.getRange(3,3,rows)); usdFmt(st.getRange(3,4,rows));
  st.getRange(3,5,rows).setNumberFormat('0.00');

  /* Device CPA flags */
  st.appendRow(['']); st.appendRow(['Device Performance Flags (30 days)']);
  st.appendRow(['Campaign','Device','Clicks','Conv','CPA','Note']);
  try{
    const dev=AdsApp.report(`
      SELECT campaign.name, segments.device,
             metrics.clicks, metrics.conversions, metrics.cost_micros
      FROM   campaign
      WHERE  campaign.advertising_channel_type = 'SEARCH'
        AND  segments.date DURING LAST_30_DAYS
        AND  metrics.cost_micros > 0`);
    const it=dev.rows(), stats={};
    while(it.hasNext()){
      const r=it.next(), c=r['campaign.name'], d=r['segments.device'];
      if(!stats[c]) stats[c]={};
      stats[c][d]={clk:+r['metrics.clicks'], conv:+r['metrics.conversions'],
                   cost:µ2$(+r['metrics.cost_micros'])};
    }
    Object.keys(stats).forEach(c=>{
      const m=stats[c]['MOBILE'] || {clk:0,conv:0,cost:0};
      const d=stats[c]['DESKTOP']|| {clk:0,conv:0,cost:0};
      const cpaM=m.conv?m.cost/m.conv:'', cpaD=d.conv?d.cost/d.conv:'';
      if(cpaM && cpaD){
        const diff=Math.abs(cpaM-cpaD)/Math.min(cpaM,cpaD);
        if(diff>0.3){
          const worst=cpaM>cpaD?['MOBILE',m,cpaM]:['DESKTOP',d,cpaD];
          st.appendRow([c,worst[0],worst[1].clk,worst[1].conv,
                        worst[2],'CPA 30 %+ higher – review bids / UX']);
        }
      }
    });
  }catch(e){ st.appendRow(['Device report error','','','','',e.message]); }

  /* Google-Ads Recommendations */
  st.appendRow(['']); st.appendRow(['Google Ads Recommendations (live)']);
  st.appendRow(['Type','Scope','Info']);
  try{
    const recs=AdsApp.recommendations().get();
    if(!recs.hasNext()) st.appendRow(['No recommendations','','']);
    while(recs.hasNext()){
      const r=recs.next();
      const scope=r.getCampaign()?r.getCampaign().getName():'Account';
      st.appendRow([r.getType(),scope,'(review/apply in Google Ads UI)']);
    }
  }catch(e){ st.appendRow(['Recommendations error','',''+e.message]); }
  finalise(st);
}

/*────────────────── 6. ADS ──────────────────*/
function buildAdsTab(ss){
  const sh = ss.insertSheet('Ads');
  AdsApp.report(`
    SELECT ad_group_ad.ad.id, campaign.name, ad_group.name, ad_group_ad.ad.type,
           metrics.impressions, metrics.clicks, metrics.ctr,
           metrics.conversions, metrics.average_cpc, metrics.cost_micros
    FROM   ad_group_ad
    WHERE  campaign.advertising_channel_type = 'SEARCH'
      AND  ad_group_ad.status = 'ENABLED'
      AND  segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'`)
    .exportToSheet(sh);

  const cpaCol = sh.getLastColumn() + 1, recCol = cpaCol + 1;
  sh.getRange(1, cpaCol, 1, 2).setValues([['CPA (USD)', 'Recommendation']]);

  convertMicros(sh, 9); convertMicros(sh, 10);
  usdFmt(sh.getRange(2, 9, sh.getLastRow() - 1, 2));
  pctFmt(sh.getRange(2, 7, sh.getLastRow() - 1));

  const rows = sh.getLastRow() - 1;
  const cost = sh.getRange(2, 10, rows).getValues();
  const conv = sh.getRange(2, 8,  rows).getValues();
  const CPA  = conv.map((c, i) => [c[0] > 0 ? cost[i][0] / c[0] : '']);
  sh.getRange(2, cpaCol, rows).setValues(CPA);
  usdFmt(sh.getRange(2, cpaCol, rows));

  const clk     = sh.getRange(2, 6, rows).getValues();
  const ctr     = sh.getRange(2, 7, rows).getValues();
  const clickTH = cfg('CLICK_THRESHOLD');

  for (let i = 0; i < rows; i++) {
    const clicks = clk[i][0], convs = conv[i][0], ctrP = ctr[i][0];
    let advice   = '', color = '';
    if (convs > 0) {
      if (ctrP < 0.015){ advice = 'Conv but low CTR – test copy'; color = REVIEW; }
      else              { advice = 'Performing well – scale';      color = TOP; }
    } else {
      if (clicks >= clickTH * 10) advice = 'Pause / rewrite – spend, no conv', color = BOTTOM;
      else if (ctrP < 0.01)       advice = 'Improve copy (CTR <1 %)',         color = REVIEW;
      else                        advice = 'Monitor';
    }
    sh.getRange(i + 2, recCol).setValue(advice);
    if (color) sh.getRange(i + 2, 1, 1, recCol).setBackground(color);
  }
  finalise(sh);
}

/*────────────────── 7. CAMPAIGN-KW ──────────────────*/
function buildCampaignKW(ss, spend){
  const list = AdsApp.report(`
    SELECT campaign.id, campaign.name
    FROM   campaign
    WHERE  campaign.advertising_channel_type = 'SEARCH'
      AND  segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'`).rows();
  const used={};
  while(list.hasNext()){
    const r=list.next(), id=r['campaign.id'];
    if(!spend[id]||spend[id]<=0) continue;

    let base=r['campaign.name'].substring(0,10)||'Campaign';
    let tab=base, n=1; while(used[tab]) tab=`${base}-${n++}`;
    used[tab]=true; const sh=ss.insertSheet(tab);

    AdsApp.report(`
      SELECT ad_group_criterion.keyword.text,
             ad_group_criterion.keyword.match_type,
             metrics.clicks, metrics.impressions, metrics.ctr,
             metrics.conversions, metrics.conversions_value,
             metrics.average_cpc, metrics.cost_micros,
             ad_group_criterion.quality_info.quality_score
      FROM   keyword_view
      WHERE  campaign.id = ${id}
        AND  segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'`)
      .exportToSheet(sh);
    if(sh.getLastRow()<2){ ss.deleteSheet(sh); continue; }

    const url=`https://ads.google.com/aw/campaigns/detail?campaignId=${id}`;
    sh.getRange('A1').setFormula(`=HYPERLINK("${url}","Open in Google Ads")`);

    const cpaCol = sh.getLastColumn() + 1,
          roasCol = cpaCol + 1,
          actCol  = roasCol + 1;
    sh.getRange(1, cpaCol, 1, 3)
      .setValues([['CPA (USD)', 'ROAS', 'Action']]);

    convertMicros(sh, 8); convertMicros(sh, 9);

    const rows = sh.getLastRow() - 1;
    const cost = sh.getRange(2, 9, rows).getValues();
    const conv = sh.getRange(2, 6, rows).getValues();
    const val  = sh.getRange(2, 7, rows).getValues();

    const CPA  = conv.map((c, i) => [c[0] > 0 ? cost[i][0] / c[0] : '']);
    const ROAS = val .map((v, i) => [cost[i][0] > 0 ? v[0] / cost[i][0] : '']);
    sh.getRange(2, cpaCol, rows).setValues(CPA);
    sh.getRange(2, roasCol, rows).setValues(ROAS);
    usdFmt(sh.getRange(2, 8, rows, 2));
    usdFmt(sh.getRange(2, cpaCol, rows));
    pctFmt(sh.getRange(2, 5, rows));

    /* Action column & row colours */
    const clickTH = cfg('CLICK_THRESHOLD');
    const text = sh.getRange(2, 1, rows).getValues();
    const clicks = sh.getRange(2, 3, rows).getValues();
    const convs  = sh.getRange(2, 6, rows).getValues();
    for(let i=0;i<rows;i++){
      const kw = text[i][0] || '';
      const isBrand = BRAND_PATTERNS.some(p=>p.test(kw));
      const c = clicks[i][0], v = convs[i][0];
      let act='', col='';
      if(v >= 1)              { act='Keep';    col=TOP;    }
      else if(isBrand && c)   { act='Review';  col=REVIEW; }
      else if(c >= clickTH)   { act='Exclude'; col=BOTTOM; }
      if(act){
        sh.getRange(i+2, actCol).setValue(act);
        sh.getRange(i+2, 1, 1, actCol).setBackground(col);
      }
    }

    /* Top / Bottom-10 block */
    const full=sh.getRange(2, 1, rows, actCol).getValues()
      .map(row=>({ctr:row[4]||0, cpa:row[cpaCol-1]||999999, data:row}));
    full.sort((a,b)=> (b.ctr!==a.ctr) ? b.ctr-a.ctr : a.cpa-b.cpa);
    const top10    = full.slice(0,10).map(o=>o.data);
    const bottom10 = full.slice(-10).map(o=>o.data);

    let s=rows+3;
    sh.getRange(s,1).setValue('▲ TOP 10 Keywords');
    sh.getRange(s+1,1,top10.length,actCol).setValues(top10)
      .setBackground(TOP);

    s+=top10.length+2;
    sh.getRange(s,1).setValue('▼ BOTTOM 10 Keywords');
    sh.getRange(s+1,1,bottom10.length,actCol).setValues(bottom10)
      .setBackground(BOTTOM);

    finalise(sh);
  }
}

/*────────────────── 8. SEARCH TERMS ──────────────────*/
function buildSearchTerms(ss){
  const sh = ss.insertSheet('Search Terms');
  AdsApp.report(`
    SELECT  search_term_view.search_term, search_term_view.status,
            campaign.name, ad_group.name,
            segments.keyword.info.text, segments.keyword.info.match_type,
            segments.search_term_match_type,
            metrics.clicks, metrics.impressions, metrics.ctr,
            metrics.conversions, metrics.average_cpc, metrics.cost_micros
    FROM    search_term_view
    WHERE   segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
      AND   metrics.cost_micros > 0
      AND   campaign.advertising_channel_type = 'SEARCH'`)
    .exportToSheet(sh);

  const hdr=['Search Term','ST Status','Campaign','Ad Group','Keyword',
             'KW Match','ST Match','Clicks','Impr','CTR','Conv',
             'Avg CPC','Cost','CPA','Brand?','Action'];
  sh.getRange(1,1,1,hdr.length).setValues([hdr]);

  convertMicros(sh, 12); convertMicros(sh, 13);
  usdFmt(sh.getRange(2, 12, sh.getLastRow() - 1, 2));
  pctFmt(sh.getRange(2, 10, sh.getLastRow() - 1));

  const rows = sh.getLastRow()-1;
  if(rows<1){ finalise(sh); return; }

  const cost=sh.getRange(2,13,rows).getValues();
  const conv=sh.getRange(2,11,rows).getValues();
  const CPA = conv.map((c,i)=>[ c[0]>0 ? cost[i][0]/c[0] : '' ]);
  sh.getRange(2,14,rows).setValues(CPA);

  const terms = sh.getRange(2, 1, rows).getValues();
  const clicks= sh.getRange(2, 8, rows).getValues();
  const clickTH=cfg('CLICK_THRESHOLD');

  const flags=[], acts=[];
  for(let i=0;i<rows;i++){
    const txt=terms[i][0]||'';
    const isBrand=BRAND_PATTERNS.some(p=>p.test(txt));
    const c=clicks[i][0], v=conv[i][0];
    let act='';
    if(v>=1)             act='Add';
    else if(isBrand && c)act='Review';
    else if(c>=clickTH)  act='Exclude';
    flags.push([isBrand]);
    acts .push([act]);
  }
  sh.getRange(2,15,rows).setValues(flags);
  sh.getRange(2,16,rows).setValues(acts);

  acts.forEach((a,i)=>{
    const row=sh.getRange(i+2,1,1,16);
    if(a[0]==='Add')     row.setBackground(TOP);
    if(a[0]==='Exclude') row.setBackground(BOTTOM);
    if(a[0]==='Review')  row.setBackground(REVIEW);
  });
  finalise(sh);
}