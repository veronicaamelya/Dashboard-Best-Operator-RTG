// =====================================================
// Dashboard Nominasi Operator RTG – TPS
// dashboard.js
// =====================================================

// ═══════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════
let wb = null;           // loaded XLSX workbook
let sheets = {};         // { key → actual sheet name | null }
let allData = {};        // { periode → { groups, stats, bobot } }
let curPeriode = null;

// Sheet detection config
const SCFG = {
  raw:     { label:'RAW DATA',            kw:['RAW DATA','RAWDATA'],              req:true,  ico:'📊' },
  manning: { label:'Manning Operator RTG',kw:['MANNING OPERATOR','MANNING OP'],  req:true,  ico:'👥' },
  presensi:{ label:'Presensi',            kw:['PRESENSI','ABSENSI'],             req:false, ico:'📋' },
  k3:      { label:'K3',                  kw:['K3','INSIDEN','INCIDENT'],        req:false, ico:'⚠️' },
  hsc:     { label:'HSC RF 133',          kw:['HSC','HOT SEAT'],                 req:false, ico:'📦' },
  match:   { label:'Match RF 60',         kw:['MATCH','RF_60','RF60'],           req:false, ico:'🎯' },
  kess:    { label:'Kesesuaian Manning',  kw:['KESESUAIAN'],                     req:false, ico:'🔧' },
};

// RTG brand by unit number
const brand = n => {
  const v = parseInt(String(n).replace(/\D/g,''));
  if (v>=38 && v<=53) return 'kone';
  if (v>=54 && v<=59) return 'kalmar';
  if (v>=60 && v<=74) return 'dinson';
  return null;
};

// ═══════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════
const toN = v => { const n=parseFloat(String(v??'').replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n; };
const norm = s => String(s||'').replace(/^\d+\s*[-–]\s*/,'').trim().toUpperCase().replace(/\s+/g,' ');

function findSheet(keywords) {
  if (!wb) return null;
  for (const kw of keywords) {
    const sn = wb.SheetNames.find(s =>
      s.toUpperCase().replace(/[\s_\-]/g,'').includes(kw.toUpperCase().replace(/[\s_\-]/g,''))
    );
    if (sn) return sn;
  }
  return null;
}

function getRows(key) {
  const sn = sheets[key]; if (!sn || !wb) return null;
  return XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, defval:null, raw:false});
}

// ═══════════════════════════════════════════════════
// BOBOT (permanen — sesuai formula RAW DATA Excel)
// ═══════════════════════════════════════════════════
const getB = () => ({ D:35, K:20, Kal:20, M:15, Man:10 });

// ═══════════════════════════════════════════════════
// FILE UPLOAD
// ═══════════════════════════════════════════════════
function handleDrop(e) { const f=e.dataTransfer.files[0]; if(f) loadFile(f); }

function resetFile() {
  wb=null; sheets={};
  document.getElementById('uz').className='upload-zone';
  document.getElementById('uz').onclick=()=>document.getElementById('fi').click();
  document.getElementById('uz-ico').textContent='📂';
  document.getElementById('uz-title').textContent='Drag & drop file Excel ke sini';
  document.getElementById('uz-desc').textContent='1 file berisi semua sheet: RAW DATA, Manning, Presensi, K3, dll.';
  document.getElementById('uz-chk').innerHTML='';
  document.getElementById('uz-extra').textContent='';
  document.getElementById('fi').value='';
  updStatus();
}

function loadFile(file) {
  if (!file) return;
  const rd = new FileReader();
  rd.onload = e => {
    try {
      wb = XLSX.read(e.target.result, {type:'array', cellDates:true});
      sheets = {};
      for (const [k,c] of Object.entries(SCFG)) sheets[k] = findSheet(c.kw);
      renderChecklist(file.name, wb.SheetNames);
      updStatus();
    } catch(err) {
      document.getElementById('uz').className='upload-zone err';
      document.getElementById('uz-title').textContent='❌ Gagal membaca file';
      document.getElementById('uz-desc').textContent=err.message;
    }
  };
  rd.readAsArrayBuffer(file);
}

function renderChecklist(fname, sheetNames) {
  const uz = document.getElementById('uz');
  uz.className = 'upload-zone ok';
  uz.onclick = null;
  document.getElementById('uz-ico').textContent = '✅';
  document.getElementById('uz-title').textContent = fname;
  document.getElementById('uz-desc').textContent = sheetNames.length + ' sheet ditemukan';

  let h = '<div class="chk-grid">';
  for (const [k,c] of Object.entries(SCFG)) {
    const found = sheets[k];
    const cls = found ? 'found' : (c.req ? 'miss' : 'opt');
    h += `<div class="chk-item ${cls}"><div class="chk-dot"></div>
      <span>${c.ico} ${c.label}</span>
      <span style="margin-left:auto;font-size:9px">${found ? '→ '+found : (c.req ? '❌ WAJIB' : 'opsional')}</span>
    </div>`;
  }
  h += '</div>';
  document.getElementById('uz-chk').innerHTML = h;

  const used = new Set(Object.values(sheets).filter(Boolean));
  const unk = sheetNames.filter(s=>!used.has(s));
  document.getElementById('uz-extra').textContent =
    unk.length ? 'Sheet tidak dikenali: '+unk.join(', ') : '';
}

function updStatus() {
  const el=document.getElementById('pstat'), btn=document.getElementById('go-btn');
  if (!wb) { el.textContent='Upload file Excel terlebih dahulu.'; btn.disabled=true; return; }
  const miss = Object.entries(SCFG).filter(([k,c])=>c.req&&!sheets[k]).map(([,c])=>c.label);
  if (miss.length) { el.textContent='⚠️ Sheet wajib tidak ditemukan: '+miss.join(', '); btn.disabled=true; }
  else {
    const n = Object.values(sheets).filter(Boolean).length;
    el.textContent = `✅ Siap! ${n}/${Object.keys(SCFG).length} sheet terdeteksi. Klik Proses.`;
    btn.disabled = false;
  }
}

// ═══════════════════════════════════════════════════
// PARSERS
// ═══════════════════════════════════════════════════

// RAW DATA: col0=No, col1=PIN, col2=Nama, col3=NIPP, col4=Group,
// col5=Dinson, col6=Kone, col7=Kalmar,
// col9=MatchPlan, col10=NotMatchPlan,
// col12=MatchMan, col13=NotMatchMan, col15=PointTotal
function parseRaw() {
  const rows = getRows('raw'); if (!rows) return [];
  const ops = [];
  for (let i=1; i<rows.length; i++) {
    const r = rows[i]; if (!r || !r[1]) continue;
    const pin = String(r[1]).trim().replace(/\D/g,'');
    if (!pin) continue;
    const grup = String(r[4]||'').trim().toUpperCase();
    if (!['A','B','C','D'].includes(grup)) continue;
    const namaRaw = String(r[2]||'').trim();
    const nama = namaRaw.replace(/^\d+\s*[-–]\s*/,'').trim();
    ops.push({
      no: toN(r[0]), pin, nama, nipp: String(r[3]||'').trim(), group: grup,
      mvDinson: toN(r[5]), mvKone: toN(r[6]), mvKalmar: toN(r[7]),
      matchPlan: toN(r[9]), notMatchPlan: toN(r[10]),
      matchMan: toN(r[12]), notMatchMan: toN(r[13]),
      rawPoint: toN(r[15]),
    });
  }
  return ops;
}

// MANNING: R1=header(NO|NAMA|NIPP|GRUP), R2+=data
function parseManning() {
  const rows = getRows('manning'); if (!rows) return {};
  const map = {};
  let hdr = 0;
  for (let i=0; i<Math.min(rows.length,5); i++) {
    const r = rows[i]; if (!r) continue;
    const s = r.map(v=>String(v||'')).join('|').toUpperCase();
    if (s.includes('NAMA') && s.includes('NIPP') && s.includes('GRUP')) { hdr=i; break; }
  }
  for (let i=hdr+1; i<rows.length; i++) {
    const r = rows[i]; if (!r || !r[0]) continue;
    const namaRaw = String(r[1]||'').trim();
    const pm = namaRaw.match(/^(\d{4})\s*[-–]/);
    if (!pm) continue;
    const gs = String(r[3]||'').trim().toUpperCase();
    if (!['A','B','C','D'].includes(gs)) continue;
    const pin = pm[1];
    map[pin] = {
      pin, nama: namaRaw.replace(/^\d+\s*[-–]\s*/,'').trim(),
      nipp: String(r[2]||'').trim(), group: gs,
    };
  }
  return map;
}

// PRESENSI: col2=NIPP, col14=TanpaKet, col15=Terlambat, col16=PC, col17=TdkMasuk, col18=TdkKeluar
function parsePresensi() {
  const rows = getRows('presensi'); if (!rows) return {};
  const map = {};
  let hdr = 2;
  for (let i=0; i<Math.min(rows.length,5); i++) {
    const r = rows[i]; if (!r) continue;
    const s = r.map(v=>String(v||'')).join('|').toUpperCase();
    if (s.includes('NIPP') && s.includes('TERLAMBAT')) { hdr=i; break; }
  }
  const h = rows[hdr] || [];
  let cNipp=2, cTK=14, cTel=15, cPC=16, cTM=17, cTK2=18;
  h.forEach((v,i)=>{
    const s = String(v||'').toUpperCase().replace(/[\s\n]+/g,'');
    if (s==='NIPP') cNipp=i;
    if (s.includes('TANPAKETERANGAN')) cTK=i;
    if (s.includes('TERLAMBATMASUK')||s==='TERLAMBAT') cTel=i;
    if (s.includes('PULANGCEPAT')) cPC=i;
    if (s.includes('TIDAKABSENMASUK')) cTM=i;
    if (s.includes('TIDAKABSENKELUAR')||s.includes('TIDAKABSENPULANG')) cTK2=i;
  });
  for (let i=hdr+1; i<rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const nipp = String(r[cNipp]||'').trim().replace(/\D/g,'');
    if (!nipp || nipp.length<5) continue;
    map[nipp] = {
      tanpaKet: toN(r[cTK]), telat: toN(r[cTel]),
      pc: toN(r[cPC]), tdkMasuk: toN(r[cTM]), tdkKeluar: toN(r[cTK2]),
    };
  }
  return map;
}

// parseK3(manMap) → Set of PIN string
// Alur: baca nama operator RTG dari K3 → lookup ke Manning → ambil PIN
// Matching: exact → 2-kata → trigram 4-char per kata
// Support Format A (Equipment col) dan Format B (Nama + Jabatan col / Incident Report)
function parseK3(manMap) {
  const rows = getRows('k3'); if (!rows) return new Set();

  // Bangun reverse lookup: norm(nama) → pin dari Manning
  const nameToPinMap = {};
  for (const [pin, info] of Object.entries(manMap)) {
    nameToPinMap[norm(info.nama)] = pin;
  }

  // Helper: nama K3 → PIN via Manning (3-level matching)
  function lookupPin(namaRaw) {
    const key = norm(namaRaw);
    if (nameToPinMap[key]) return nameToPinMap[key];

    const wA = key.split(' ');
    for (const [mKey, pin] of Object.entries(nameToPinMap)) {
      const wB = mKey.split(' ');
      // Level 2: 2-kata pertama exact
      const w2A = wA.slice(0,2).join(' '), w2B = wB.slice(0,2).join(' ');
      if (w2A.length > 3 && w2A === w2B) return pin;
      // Level 3: trigram 4-char per kata (handle RACHMAT vs RACHMAD, dll)
      if (wA.length >= 2 && wB.length >= 2) {
        const t = s => s.slice(0,4);
        if (t(wA[0]) === t(wB[0]) && t(wA[1]) === t(wB[1])) return pin;
      }
    }
    return null;
  }

  // Deteksi format sheet K3
  let formatB = false;
  let cNama = 7, cJabatan = 8, cKronologi = 6;
  let cEquip = 11, cOp = 12;

  for (let i = 0; i < Math.min(rows.length, 8); i++) {
    const r = rows[i]; if (!r) continue;
    const joined = r.map(v => String(v||'').toUpperCase().replace(/\s+/g,'')).join('|');
    if (joined.includes('NAMA') && joined.includes('JABATAN')) {
      formatB = true;
      r.forEach((v, ci) => {
        const s = String(v||'').toUpperCase().replace(/\s+/g,'');
        if (s === 'NAMA')            cNama      = ci;
        if (s === 'JABATAN')         cJabatan   = ci;
        if (s.includes('KRONOLOGI')) cKronologi = ci;
      });
      break;
    }
    if (joined.includes('EQUIP')) {
      r.forEach((v, ci) => {
        const s = String(v||'').toUpperCase().replace(/\s+/g,'');
        if ((s==='EQUIPMENT'||s.includes('EQUIP')) && ci>=8 && ci<=14) cEquip = ci;
        if ((s.includes('NAMAEQUIP')||s.includes('OPERATORN')) && ci>=9) cOp = ci;
      });
      break;
    }
  }

  const k3Pins = new Set();

  for (const r of rows) {
    if (!r || !r[0]) continue;
    const no = parseInt(String(r[0]||'').replace(/\D/g,'')); if (isNaN(no) || no <= 0) continue;

    let namaRaw = '';
    if (formatB) {
      const jabatan   = String(r[cJabatan]   || '').toUpperCase();
      const kronologi = String(r[cKronologi] || '').toUpperCase();
      namaRaw         = String(r[cNama]      || '').trim();
      if (!namaRaw) continue;
      const isRTGOp = jabatan.includes('RTG') ||
                      (kronologi.includes('RTG') && jabatan.includes('OPERATOR'));
      if (!isRTGOp) continue;
    } else {
      const equip = String(r[cEquip] || '').toUpperCase();
      if (!equip.includes('RTG')) continue;
      namaRaw = String(r[cOp] || '').trim();
      if (!namaRaw) continue;
    }

    const pin = lookupPin(namaRaw);
    if (pin) {
      k3Pins.add(pin);
    } else {
      // Tidak ketemu di Manning — simpan nama sebagai fallback dengan prefix '#'
      k3Pins.add('#' + norm(namaRaw));
    }
  }

  return k3Pins;
}

// ═══════════════════════════════════════════════════
// MAIN PROCESS
// ═══════════════════════════════════════════════════
function processAll() {
  const periode = document.getElementById('periode-in').value.trim() ||
    ('PERIODE '+new Date().toLocaleDateString('id-ID',{month:'long',year:'numeric'}).toUpperCase());
  if (!wb) { showAlert('err','❌ Belum ada file!'); return; }
  if (!sheets.raw)     { showAlert('err','❌ Sheet RAW DATA tidak ditemukan!'); return; }
  if (!sheets.manning) { showAlert('err','❌ Sheet Manning tidak ditemukan!'); return; }
  const b = getB();

  loading('Membaca RAW DATA...');
  setTimeout(()=>{
    try {
      const rawOps = parseRaw();
      if (!rawOps.length) { hide(); showAlert('err','❌ RAW DATA kosong atau format tidak sesuai.'); return; }

      loading('Membaca Manning...');
      const manMap = parseManning();

      loading('Membaca Presensi & K3...');
      const presMap = sheets.presensi ? parsePresensi() : {};
      const k3Pins  = sheets.k3       ? parseK3(manMap) : new Set();

      loading('Menghitung poin...');

      const groups = {A:[],B:[],C:[],D:[]};
      for (const op of rawOps) {
        const manInfo = manMap[op.pin];
        const nipp = manInfo?.nipp || op.nipp || '';
        const presData = nipp ? presMap[nipp] : null;

        // K3 check by PIN — akurat, tidak bergantung ejaan nama
        const k3count = k3Pins.has(op.pin) ? 1
                      : k3Pins.has('#' + norm(op.nama)) ? 1   // fallback nama
                      : 0;

        const g = op.group;
        if (!groups[g]) continue;
        groups[g].push({...op, nipp, presData, k3count});
      }

      // Hitung avg moves per brand per grup, lalu score
      for (const g of ['A','B','C','D']) {
        const ops = groups[g]; if (!ops.length) continue;

        // Avg dihitung dari SEMUA operator (termasuk yang 0 moves),
        // sesuai formula di RAW DATA Excel
        const avg = field => ops.length ? ops.reduce((s,o)=>s+o[field],0)/ops.length : 0;
        const avgD=avg('mvDinson'), avgK=avg('mvKone'), avgKal=avg('mvKalmar');

        for (const op of ops) {
          const pD   = avgD   > 0 ? b.D   * Math.min(1, op.mvDinson/avgD)   : 0;
          const pK   = avgK   > 0 ? b.K   * Math.min(1, op.mvKone/avgK)     : 0;
          const pKal = avgKal > 0 ? b.Kal * Math.min(1, op.mvKalmar/avgKal) : 0;

          const totMP = op.matchPlan + op.notMatchPlan;
          const pM = totMP > 0 ? b.M * (op.matchPlan/totMP) : 0;

          const totMM = op.matchMan + op.notMatchMan;
          const pMan = totMM > 0 ? b.Man * (op.matchMan/totMM) : 0;

          const totalPoin = pD + pK + pKal + pM + pMan;

          const pr = op.presData;
          const absenViol = pr
            ? (pr.tanpaKet>0||pr.telat>0||pr.pc>0||pr.tdkMasuk>0||pr.tdkKeluar>0)
            : false;
          const eligible = !absenViol && op.k3count===0;

          Object.assign(op, {
            avgD, avgK, avgKal,
            pD, pK, pKal, pM, pMan, totalPoin,
            absenViol, eligible,
            tanpaKet:  pr?.tanpaKet  || 0,
            telat:     pr?.telat     || 0,
            pc:        pr?.pc        || 0,
            tdkMasuk:  pr?.tdkMasuk  || 0,
            tdkKeluar: pr?.tdkKeluar || 0,
            hasPres:   !!pr,
          });
        }
        ops.sort((a,z)=>z.totalPoin-a.totalPoin);
      }

      let totalOps=0, totalMoves=0, totalK3=0;
      for (const g of ['A','B','C','D']) {
        groups[g].forEach(o=>{ totalOps++; totalMoves+=o.mvDinson+o.mvKone+o.mvKalmar; if(o.k3count) totalK3++; });
      }

      allData[periode] = {
        periode, uploadedAt: new Date().toISOString(),
        groups, bobot:{...b},
        stats:{ totalOps, totalMoves, totalK3,
          eligible: ['A','B','C','D'].flatMap(g=>groups[g]).filter(o=>o.eligible).length },
        sheetsFound: Object.entries(sheets).filter(([,v])=>v).map(([k])=>k),
      };
      curPeriode = periode;

      hide();
      showAlert('suc', `✅ <strong>${periode}</strong> berhasil! ${totalOps} operator, ${['A','B','C','D'].filter(g=>groups[g].length).length} grup.`);
      const _pb1=document.getElementById('periodeBadge'); if(_pb1) _pb1.textContent=periode;
      sw('dash', document.querySelectorAll('.tab')[1]);

    } catch(err) {
      hide();
      showAlert('err','❌ Error: '+err.message);
      console.error(err);
    }
  }, 80);
}

// ═══════════════════════════════════════════════════
// TAB SWITCH
// ═══════════════════════════════════════════════════
function sw(name, btn) {
  document.querySelectorAll('.pane').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(b=>b.classList.remove('on'));
  document.getElementById('pane-'+name).classList.add('on');
  if (btn) btn.classList.add('on');
  if (name==='dash') rDash();
  if (name==='tbl')  rTbl();
  if (name==='det')  rDet();
  if (name==='thn')  rThn();
}

// ═══════════════════════════════════════════════════
// ALERT & LOADING
// ═══════════════════════════════════════════════════
function showAlert(t,m){const e=document.getElementById('alertBox');e.className='alert on '+t;e.innerHTML=m;setTimeout(()=>e.classList.remove('on'),9000);}
function loading(t){document.getElementById('ovt').textContent=t||'Memproses...';document.getElementById('ov').classList.add('on');}
function hide(){document.getElementById('ov').classList.remove('on');}

// ═══════════════════════════════════════════════════
// RENDER DASHBOARD
// ═══════════════════════════════════════════════════
function rDash() {
  const keys = Object.keys(allData);
  if (!keys.length) { document.getElementById('dash-content').innerHTML='<div class="empty"><div class="ei">📊</div><h3>Belum ada data</h3></div>'; return; }
  if (!curPeriode||!allData[curPeriode]) curPeriode=keys[keys.length-1];
  const pd = allData[curPeriode];
  const {groups,stats,bobot} = pd;
  const _pb2=document.getElementById('periodeBadge'); if(_pb2) _pb2.textContent=curPeriode;

  const chips = keys.map(p=>`<div class="chip ${p===curPeriode?'on':''}" onclick="curPeriode='${p}';rDash()">
    📅 ${p}<span class="chip-x" onclick="event.stopPropagation();delP('${p}')"> ✕</span></div>`).join('');

  const statHtml = `<div class="stats">
    <div class="sc navy"><div class="sc-ico">👤</div><div><div class="sc-lbl">Total Operator</div><div class="sc-val">${stats.totalOps}</div><div class="sc-sub">${['A','B','C','D'].filter(g=>groups[g]?.length).length} grup aktif</div></div></div>
    <div class="sc gold"><div class="sc-ico">📦</div><div><div class="sc-lbl">Total Moves</div><div class="sc-val">${stats.totalMoves.toLocaleString('id')}</div><div class="sc-sub">Dinson+Kone+Kalmar</div></div></div>
    <div class="sc green"><div class="sc-ico">✅</div><div><div class="sc-lbl">Memenuhi Syarat</div><div class="sc-val">${stats.eligible}</div><div class="sc-sub">Bisa dinominasikan</div></div></div>
    <div class="sc red"><div class="sc-ico">🚫</div><div><div class="sc-lbl">Gugur</div><div class="sc-val">${stats.totalOps-stats.eligible}</div><div class="sc-sub">Presensi / K3</div></div></div>
    <div class="sc teal"><div class="sc-ico">⚠️</div><div><div class="sc-lbl">Terlibat K3</div><div class="sc-val">${stats.totalK3}</div><div class="sc-sub">Insiden RTG</div></div></div>
  </div>`;

  const gHtml = `<div class="grids">${['A','B','C','D'].map(g=>{
    const ops=groups[g]||[]; if(!ops.length) return '';
    const winner=ops.find(o=>o.eligible);
    const maxPt=ops[0]?.totalPoin||100;
    const medals=['🥇','🥈','🥉'];
    const elig=ops.filter(o=>o.eligible).length;
    const rows=ops.slice(0,10).map((op,i)=>{
      const rc=i===0?'r1':i===1?'r2':i===2?'r3':'';
      const pct=maxPt>0?Math.round(op.totalPoin/maxPt*100):0;
      const why=op.k3count>0?'K3':(op.absenViol?'Absen':'');
      return `<div class="op ${rc} ${op.eligible?'':'disq'}" onclick="showDet('${curPeriode}','${op.pin}')">
        <div class="rdot">${i<3?medals[i]:i+1}</div>
        <div class="op-inf">
          <div class="op-name">${op.nama}</div>
          <div class="op-meta">PIN ${op.pin} · ${(op.mvDinson+op.mvKone+op.mvKalmar).toLocaleString('id')} moves</div>
        </div>
        ${op.eligible?`<span class="tag tag-ok">✅</span>`:`<span class="tag tag-no">🚫${why}</span>`}
        <div class="mini-bar"><div class="mb-bg"><div class="mb-fill" style="width:${pct}%"></div></div></div>
        <div class="op-pt ${i===0&&op.eligible?'hi':''}">${op.totalPoin.toFixed(1)}</div>
      </div>`;
    }).join('');
    return `<div class="gc">
      <div class="gc-hdr">
        <div><h3>GRUP ${g}</h3><div class="gc-sub">${ops.length} operator · ${elig} memenuhi syarat</div></div>
        <div class="gc-win">🏆 ${winner?.nama.split(' ').slice(0,2).join(' ')||'—'}</div>
      </div>${rows}
    </div>`;
  }).join('')}</div>`;

  const wRows=['A','B','C','D'].map(g=>{
    const winner=(groups[g]||[]).find(o=>o.eligible); if(!winner) return '';
    return `<tr class="tr-win">
      <td><span class="tag tag-${g}">GRUP ${g}</span></td>
      <td>${winner.pin}</td><td><strong>${winner.nama}</strong></td>
      <td>${winner.mvDinson.toLocaleString('id')}</td><td>${winner.mvKone.toLocaleString('id')}</td><td>${winner.mvKalmar.toLocaleString('id')}</td>
      <td>${winner.pD.toFixed(2)}</td><td>${winner.pK.toFixed(2)}</td><td>${winner.pKal.toFixed(2)}</td>
      <td>${winner.pM.toFixed(2)}</td><td>${winner.pMan.toFixed(2)}</td>
      <td><strong>${winner.totalPoin.toFixed(2)}</strong></td>
      <td><span class="tag tag-win">🏆 NOMINASI</span></td>
    </tr>`;
  }).join('');

  document.getElementById('dash-content').innerHTML=`
    <div style="margin-bottom:16px"><div class="stitle"><h2>📅 Periode</h2></div><div class="chips">${chips}</div></div>
    ${statHtml}
    <div class="stitle"><h2>📊 Ranking per Grup</h2><div class="sbadge">${curPeriode}</div></div>
    <p style="font-size:11px;color:var(--muted);margin-bottom:12px">Klik nama operator untuk detail · ✅ = Memenuhi syarat · 🚫 = Gugur</p>
    ${gHtml}
    <div class="stitle"><h2>🏅 Nominasi ${curPeriode}</h2></div>
    <div class="twrap"><table>
      <thead><tr><th>Grup</th><th>PIN</th><th>Nama</th><th>Mv Dinson</th><th>Mv Kone</th><th>Mv Kalmar</th>
        <th>Poin D</th><th>Poin K</th><th>Poin Kal</th><th>Poin Match</th><th>Poin Manning</th><th>TOTAL</th><th>Status</th></tr></thead>
      <tbody>${wRows||'<tr><td colspan="13" style="text-align:center;padding:20px;color:var(--muted)">Tidak ada operator yang memenuhi syarat</td></tr>'}</tbody>
    </table></div>`;
}

function delP(p){if(!confirm(`Hapus data "${p}"?`))return;delete allData[p];if(curPeriode===p)curPeriode=Object.keys(allData).slice(-1)[0]||null;rDash();}

// ═══════════════════════════════════════════════════
// RENDER DATA LENGKAP
// ═══════════════════════════════════════════════════
function rTbl() {
  const keys=Object.keys(allData);
  if(!keys.length){document.getElementById('tbl-content').innerHTML='<div class="empty"><div class="ei">📋</div><h3>Belum ada data</h3></div>';return;}
  document.getElementById('tbl-content').innerHTML=`
    <div class="sbar">
      <input class="sinput" id="ts" placeholder="🔍 Cari nama atau PIN..." oninput="fTbl()">
      <select class="ssel" id="tg" onchange="fTbl()"><option value="">Semua Grup</option><option>A</option><option>B</option><option>C</option><option>D</option></select>
      <select class="ssel" id="tp" onchange="fTbl()">${keys.sort().map(p=>`<option ${p===curPeriode?'selected':''}>${p}</option>`).join('')}</select>
      <select class="ssel" id="tf" onchange="fTbl()"><option value="">Semua</option><option value="ok">Memenuhi Syarat</option><option value="no">Gugur</option></select>
    </div>
    <div id="tbl-rows"></div>`;
  fTbl();
}

function fTbl(){
  const q=(document.getElementById('ts')?.value||'').toUpperCase();
  const gf=document.getElementById('tg')?.value||'';
  const pf=document.getElementById('tp')?.value||curPeriode||'';
  const ff=document.getElementById('tf')?.value||'';
  const pd=allData[pf]; if(!pd) return;
  let rows='';
  for(const g of ['A','B','C','D']){
    if(gf&&gf!==g) continue;
    (pd.groups[g]||[]).forEach((op,i)=>{
      if(q&&!op.nama.toUpperCase().includes(q)&&!op.pin.includes(q)) return;
      if(ff==='ok'&&!op.eligible) return;
      if(ff==='no'&&op.eligible) return;
      const totMv=(op.mvDinson+op.mvKone+op.mvKalmar).toLocaleString('id');
      const pctMP=op.matchPlan+op.notMatchPlan>0?((op.matchPlan/(op.matchPlan+op.notMatchPlan))*100).toFixed(1)+'%':'—';
      const pctMM=op.matchMan+op.notMatchMan>0?((op.matchMan/(op.matchMan+op.notMatchMan))*100).toFixed(1)+'%':'—';
      const why=!op.eligible?(op.k3count>0?'K3':(op.absenViol?'Absen':'')):'' ;
      rows+=`<tr class="${i===0&&op.eligible?'tr-win':''} ${op.eligible?'':'tr-dq'}" style="cursor:pointer" onclick="showDet('${pf}','${op.pin}')">
        <td>${i+1}</td><td><span class="tag tag-${g}">GRUP ${g}</span></td>
        <td>${op.pin}</td><td>${op.nipp||'—'}</td><td><strong>${op.nama}</strong></td>
        <td>${op.mvDinson.toLocaleString('id')}</td><td>${op.mvKone.toLocaleString('id')}</td><td>${op.mvKalmar.toLocaleString('id')}</td>
        <td>${op.pD.toFixed(1)}</td><td>${op.pK.toFixed(1)}</td><td>${op.pKal.toFixed(1)}</td>
        <td>${pctMP}</td><td>${op.pM.toFixed(2)}</td>
        <td>${pctMM}</td><td>${op.pMan.toFixed(2)}</td>
        <td><strong>${op.totalPoin.toFixed(2)}</strong></td>
        <td>${op.hasPres?(op.tanpaKet+'/'+op.telat+'/'+op.pc+'/'+op.tdkMasuk+'/'+op.tdkKeluar):'—'}</td>
        <td>${op.k3count>0?`<span class="tag tag-no">⚠️${op.k3count}</span>`:'—'}</td>
        <td>${op.eligible?'<span class="tag tag-ok">✅</span>':`<span class="tag tag-no">🚫 ${why}</span>`}</td>
      </tr>`;
    });
  }
  document.getElementById('tbl-rows').innerHTML=`<div class="twrap">
    <div class="thdr"><h3>Data Lengkap – ${pf}</h3><div class="sbadge">${pd.stats.totalOps} Operator</div><span style="font-size:11px;color:var(--muted);margin-left:6px">Klik baris untuk detail</span></div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Rank</th><th>Grup</th><th>PIN</th><th>NIPP</th><th>Nama</th>
        <th>Mv Dinson</th><th>Mv Kone</th><th>Mv Kalmar</th>
        <th>Poin D</th><th>Poin K</th><th>Poin Kal</th>
        <th>% Match</th><th>Poin Match</th>
        <th>% Manning</th><th>Poin Manning</th>
        <th>TOTAL</th><th>Absen(TK/T/PC/TM/TK)</th><th>K3</th><th>Status</th>
      </tr></thead>
      <tbody>${rows||'<tr><td colspan="19" style="text-align:center;padding:20px;color:var(--muted)">Tidak ada data</td></tr>'}</tbody>
    </table></div>
  </div>`;
}

// ═══════════════════════════════════════════════════
// RENDER DETAIL
// ═══════════════════════════════════════════════════
function rDet() {
  const keys=Object.keys(allData);
  if(!keys.length){document.getElementById('det-content').innerHTML='<div class="empty"><div class="ei">🔍</div><h3>Belum ada data</h3></div>';return;}
  const pd=allData[curPeriode||keys[keys.length-1]];
  const allOps=['A','B','C','D'].flatMap(g=>(pd.groups[g]||[]));
  document.getElementById('det-content').innerHTML=`
    <div class="sbar">
      <select class="ssel" id="dp" style="min-width:160px" onchange="rDetOp()">${keys.sort().map(p=>`<option ${p===(curPeriode||'')?'selected':''}>${p}</option>`).join('')}</select>
      <select class="ssel" id="dop" style="flex:1;min-width:200px" onchange="rDetOp()">
        <option value="">— Pilih Operator —</option>
        ${allOps.map(o=>`<option value="${o.pin}">${o.pin} – ${o.nama} (Grup ${o.group})</option>`).join('')}
      </select>
    </div>
    <div id="det-op"></div>`;
}

function showDet(periode, pin) {
  const detTab = Array.from(document.querySelectorAll('.tab')).find(b=>b.textContent.includes('Detail'));
  sw('det', detTab);
  setTimeout(function(){
    const selP=document.getElementById('dp');
    if(selP) selP.value=periode;
    const sel=document.getElementById('dop');
    if(sel){ sel.value=pin; rDetOp(); }
  },50);
}

function rDetOp(){
  const pf=document.getElementById('dp')?.value;
  const pin=document.getElementById('dop')?.value;
  if(!pf||!pin){document.getElementById('det-op').innerHTML='';return;}
  const pd=allData[pf]; if(!pd) return;
  const op=['A','B','C','D'].flatMap(g=>(pd.groups[g]||[])).find(o=>o.pin===pin);
  if(!op) return;
  const b=pd.bobot;
  const rank=(pd.groups[op.group]||[]).findIndex(o=>o.pin===pin)+1;
  const scores=[
    {ico:'🏗️',lbl:'RTG Dinson',val:op.pD,max:b.D,
     sub:`${op.mvDinson.toLocaleString('id')} moves · avg grup ${op.avgD.toFixed(0)} · ratio ${op.avgD>0?Math.min(100,op.mvDinson/op.avgD*100).toFixed(0):0}%`,col:'#1e4080'},
    {ico:'🏗️',lbl:'RTG Kone',val:op.pK,max:b.K,
     sub:`${op.mvKone.toLocaleString('id')} moves · avg grup ${op.avgK.toFixed(0)} · ratio ${op.avgK>0?Math.min(100,op.mvKone/op.avgK*100).toFixed(0):0}%`,col:'#00A896'},
    {ico:'🏗️',lbl:'RTG Kalmar',val:op.pKal,max:b.Kal,
     sub:`${op.mvKalmar.toLocaleString('id')} moves · avg grup ${op.avgKal.toFixed(0)} · ratio ${op.avgKal>0?Math.min(100,op.mvKalmar/op.avgKal*100).toFixed(0):0}%`,col:'#D4A017'},
    {ico:'🎯',lbl:'Match Penempatan',val:op.pM,max:b.M,
     sub:`${op.matchPlan} match · ${op.notMatchPlan} tidak · ${op.matchPlan+op.notMatchPlan>0?((op.matchPlan/(op.matchPlan+op.notMatchPlan))*100).toFixed(1):0}%`,col:'#7b5ea7'},
    {ico:'👷',lbl:'Match Manning',val:op.pMan,max:b.Man,
     sub:`${op.matchMan} sesuai · ${op.notMatchMan} tidak · ${op.matchMan+op.notMatchMan>0?((op.matchMan/(op.matchMan+op.notMatchMan))*100).toFixed(1):0}%`,col:'#C0392B'},
  ];
  const sCards=scores.map(s=>`<div class="scomp">
    <div class="sc2-ico">${s.ico}</div>
    <div style="flex:1">
      <div class="sc2-lbl">${s.lbl}</div>
      <div class="sc2-val">${s.val.toFixed(2)}<span style="font-size:11px;color:var(--muted)"> / ${s.max}</span></div>
      <div class="sc2-sub">${s.sub}</div>
      <div class="sc2-bar"><div class="sc2-fill" style="width:${s.max>0?Math.min(100,s.val/s.max*100).toFixed(0):0}%;background:${s.col}"></div></div>
    </div></div>`).join('');

  const absenItems=[
    ['Tanpa Keterangan',op.tanpaKet],['Terlambat',op.telat],
    ['Pulang Cepat',op.pc],['Tidak Absen Masuk',op.tdkMasuk],['Tidak Absen Keluar',op.tdkKeluar],
  ].map(([l,v])=>`<div class="abitem" style="background:${v>0?'var(--red-pale)':'var(--green-pale)'}">
    <div class="abitem-lbl">${l}</div>
    <div class="abitem-val" style="color:${v>0?'var(--red)':'var(--green)'}">${v}</div>
  </div>`).join('');

  const eligBox=op.eligible
    ?`<div class="sybox sy-ok">✅ Memenuhi semua syarat — Presensi bersih & tidak ada insiden RTG</div>`
    :`<div class="sybox sy-no">🚫 Gugur: ${op.absenViol?'Ada pelanggaran presensi':''}${op.absenViol&&op.k3count?' + ':''}${op.k3count?`Terlibat ${op.k3count} insiden RTG`:''}</div>`;

  document.getElementById('det-op').innerHTML=`<div class="det">
    <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:16px">
      <div style="font-size:46px">👤</div>
      <div>
        <h2 style="font-size:20px;font-weight:800;color:var(--navy)">${op.nama}</h2>
        <p style="font-size:13px;color:var(--muted)">PIN: ${op.pin} · NIPP: ${op.nipp||'—'} · <span class="tag tag-${op.group}">GRUP ${op.group}</span> · Rank #${rank}</p>
      </div>
      <div style="margin-left:auto;text-align:right">
        <div style="font-size:36px;font-weight:900;color:var(--teal)">${op.totalPoin.toFixed(2)}</div>
        <div style="font-size:11px;color:var(--muted)">TOTAL POIN</div>
      </div>
    </div>
    <div class="score-grid">${sCards}</div>
    ${eligBox}
    <div style="margin-top:14px">
      <h3 style="font-size:13px;font-weight:700;color:var(--navy);margin-bottom:6px">📋 Presensi ${op.hasPres?'':'(data tidak tersedia)'}</h3>
      ${op.hasPres?`<div class="abgrid">${absenItems}</div>`:'<p style="color:var(--muted);font-size:12px">Sheet Presensi tidak diupload.</p>'}
    </div>
    ${op.k3count>0?`<div style="margin-top:12px;padding:10px 14px;background:var(--red-pale);border-radius:8px;font-size:12px;color:var(--red);font-weight:600">⚠️ Terlibat ${op.k3count} insiden RTG pada K3 — otomatis gugur</div>`:''}
  </div>`;
}

// ═══════════════════════════════════════════════════
// RENDER TAHUNAN
// ═══════════════════════════════════════════════════
function rThn(){
  const keys=Object.keys(allData).sort();
  if(!keys.length){document.getElementById('thn-content').innerHTML='<div class="empty"><div class="ei">🏅</div><h3>Belum ada data</h3></div>';return;}
  const track={};
  for(const p of keys){
    const pd=allData[p];
    for(const g of ['A','B','C','D']){
      const w=(pd.groups[g]||[]).find(o=>o.eligible); if(!w) continue;
      if(!track[w.pin]) track[w.pin]={...w,wins:[],pts:[]};
      track[w.pin].wins.push(p); track[w.pin].pts.push(w.totalPoin);
    }
  }
  const sorted=Object.values(track).sort((a,z)=>z.wins.length-a.wins.length||(z.pts.reduce((s,v)=>s+v,0)/z.pts.length)-(a.pts.reduce((s,v)=>s+v,0)/a.pts.length));
  const rows=sorted.map((w,i)=>{
    const avg=(w.pts.reduce((s,v)=>s+v,0)/w.pts.length).toFixed(1);
    return `<tr><td>${i+1}</td><td><span class="tag tag-${w.group}">GRUP ${w.group}</span></td>
      <td>${w.pin}</td><td>${w.nipp||'—'}</td><td><strong>${w.nama}</strong></td>
      <td><strong style="color:var(--teal)">${w.wins.length}×</strong></td>
      <td style="font-size:11px">${w.wins.join(', ')}</td><td>${avg}</td>
      <td>${i<4?'<span class="tag tag-win">🏅</span>':''}</td></tr>`;
  }).join('');
  const byP=keys.map(p=>{
    const pd=allData[p];
    return `<tr><td><strong>${p}</strong></td>${['A','B','C','D'].map(g=>{
      const w=(pd.groups[g]||[]).find(o=>o.eligible);
      return `<td>${w?`<strong>${w.nama}</strong><br><small style="color:var(--muted)">${w.totalPoin.toFixed(1)} poin</small>`:'<small style="color:var(--muted)">—</small>'}</td>`;
    }).join('')}</tr>`;
  }).join('');
  document.getElementById('thn-content').innerHTML=`
    <div class="stitle"><h2>🏅 Rekap Nominasi Tahunan</h2><div class="sbadge">${keys.length} Periode</div></div>
    <div class="twrap"><div class="thdr"><h3>Operator Paling Sering Menang</h3></div>
      <table><thead><tr><th>Rank</th><th>Grup</th><th>PIN</th><th>NIPP</th><th>Nama</th><th>Menang</th><th>Periode</th><th>Avg Poin</th><th>Status</th></tr></thead>
      <tbody>${rows||'<tr><td colspan="9" style="text-align:center;padding:20px">Belum cukup data</td></tr>'}</tbody></table></div>
    <div class="stitle" style="margin-top:18px"><h2>📅 Pemenang per Periode</h2></div>
    <div class="twrap"><table>
      <thead><tr><th>Periode</th><th>Pemenang Grup A</th><th>Pemenang Grup B</th><th>Pemenang Grup C</th><th>Pemenang Grup D</th></tr></thead>
      <tbody>${byP}</tbody></table></div>`;
}