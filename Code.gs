/* AKTIVA SMAN 1 NGAWEN - API v6 (Google Apps Script)
   Deploy: Deploy > New deployment > Web app > Execute as: Me > Access: Anyone.
   Set Project Settings > Time zone = Asia/Jakarta. Jalankan setupDatabase() sekali. */
const CONFIG = {SPREADSHEET_ID:'', ADMIN_USER:'admin', DEFAULT_ADMIN_PASSWORD:'Admin@12345', TZ:'Asia/Jakarta', SESSION_HOURS:12};
const SHEETS = {
  guru:['Master_Guru',['ID Guru','Nama Guru','Email','Username','PasswordHash','Aktif']],
  mapel:['Master_Mapel',['ID Mapel','Mata Pelajaran','KKM','Aktif']],
  kelas:['Master_Kelas',['ID Kelas','Kelas','Tingkat','Jurusan','Aktif']],
  siswa:['Master_Siswa',['ID Siswa','NISN','Nama Siswa','ID Kelas','Kelas','Aktif']],
  jadwal:['Master_Jadwal',['ID Jadwal','ID Guru','ID Mapel','ID Kelas','Hari','Jam Mulai','Jam Selesai','Jam Pelajaran','Aktif']],
  jurnal:['Jurnal',['Timestamp','Tanggal','Hari','Jam','ID Guru','Nama Guru','ID Mapel','Mata Pelajaran','ID Kelas','Kelas','ID Jadwal','Jam Pelajaran','Materi','Keterangan']],
  presensi:['Presensi',['Timestamp','Tanggal','Hari','Jam','ID Guru','Nama Guru','ID Mapel','Mata Pelajaran','ID Kelas','Kelas','ID Jadwal','Jam Pelajaran','ID Siswa','Nama Siswa','Status']],
  penilaian:['Penilaian',['Timestamp','Tanggal','Hari','Jam','ID Kelas','Kelas','Kehadiran','Penampilan','Pakaian','Kebersihan','Keaktifan','Perilaku','Integritas','Total Bintang','Rata Rata','ID Guru','Nama Guru']],
  refleksi:['Refleksi',['Timestamp','Tanggal','Hari','Jam','ID Guru','Nama Guru','Facts','Feelings','Findings','Future','Rating Kebahagiaan']],
  nilai:['Nilai',['Timestamp','Tanggal','ID Guru','Nama Guru','ID Mapel','Mata Pelajaran','ID Kelas','Kelas','Jenis','Topik','ID Siswa','Nama Siswa','Nilai','KKM','Tuntas']],
  agenda:['Agenda',['Timestamp','ID Agenda','ID Guru','Tanggal','Judul','Jenis','Selesai']],
  piket:['Master_Piket',['ID Piket','ID Guru','Hari','Jam Mulai','Jam Selesai','Keterangan','Aktif']],
  laporpiket:['Laporan_Piket',['Timestamp','Tanggal','Hari','Jam','ID Guru','Nama Guru','Siswa Terlambat','Siswa Izin Sakit','Guru Tidak Hadir','Kejadian','Tindak Lanjut']]
};
const MASTER = {guru:{id:'ID Guru',p:'G'},mapel:{id:'ID Mapel',p:'M'},kelas:{id:'ID Kelas',p:'K'},siswa:{id:'ID Siswa',p:'S'},jadwal:{id:'ID Jadwal',p:'J'},piket:{id:'ID Piket',p:'P'}};
const HARI = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
const STATUS = ['Hadir','Terlambat','Sakit','Izin','Alpa','Dispensasi'];
const SCORES = ['kehadiran','penampilan','pakaian','kebersihan','keaktifan','perilaku','integritas'];
let MEMO = {};
const PROP = () => PropertiesService.getScriptProperties();

/* ---------- Database ---------- */
function ss_() { return CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function hdr_(sh) { return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String); }
function setupDatabase(force) {
  if (force !== true && PROP().getProperty('DB_READY') === '8') return 'Database sudah siap.';
  const ss = ss_();
  Object.keys(SHEETS).forEach(k => {
    const name = SHEETS[k][0], hd = SHEETS[k][1];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,hd.length).setValues([hd]);
    else { const cur = hdr_(sh); hd.forEach(h => { if (cur.indexOf(h) < 0) { sh.getRange(1, cur.length+1).setValue(h); cur.push(h); } }); }
    sh.setFrozenRows(1);
    hdr_(sh).forEach((h,i) => { if (MASTER[k] || ['Tanggal','Jam','Jam Pelajaran','ID Jadwal'].indexOf(h) >= 0) sh.getRange(2,i+1,Math.max(sh.getMaxRows()-1,1),1).setNumberFormat('@'); });
  });
  adminHash_();
  PROP().setProperty('DB_READY','8');
  return 'Database siap: ' + ss.getUrl();
}
function sheet_(k) {
  const sh = ss_().getSheetByName(SHEETS[k][0]);
  if (sh) return sh;
  setupDatabase(true); return ss_().getSheetByName(SHEETS[k][0]);
}
const fmtD_ = d => Utilities.formatDate(d, CONFIG.TZ, 'dd/MM/yyyy');
function cv_(h,v) {
  if (!(v instanceof Date)) return v;
  if (h === 'Tanggal') return fmtD_(v);
  if (h === 'Timestamp') return v.toISOString();
  return Utilities.formatDate(v, CONFIG.TZ, 'HH:mm');
}
function rows_(k) {
  if (MEMO[k]) return MEMO[k];
  const v = sheet_(k).getDataRange().getValues(), h = v.shift().map(String);
  return MEMO[k] = v.filter(r => r.some(x => x !== '')).map(r => { const o = {}; h.forEach((x,i) => o[x] = cv_(x,r[i])); return o; });
}
function append_(k, objs) {
  if (!objs.length) return;
  const sh = sheet_(k), h = hdr_(sh);
  sh.getRange(sh.getLastRow()+1,1,objs.length,h.length).setValues(objs.map(o => h.map(x => o[x] === undefined ? '' : o[x])));
  delete MEMO[k];
}
function removeWhere_(k, fn) {
  const sh = sheet_(k), v = sh.getDataRange().getValues(), h = v[0].map(String);
  for (let i = v.length-1; i >= 1; i--) {
    const o = {}; h.forEach((x,j) => o[x] = cv_(x,v[i][j]));
    if (fn(o)) sh.deleteRow(i+1);
  }
  delete MEMO[k];
}
const off_ = x => String(x['Aktif']).toLowerCase() === 'false';
const by_ = (arr,key) => { const m = {}; arr.forEach(x => m[String(x[key])] = x); return m; };

/* ---------- Tanggal ---------- */
function parseD_(v) {
  if (v instanceof Date) return v;
  const s = String(v||''); let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(+m[3],+m[2]-1,+m[1],12);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1],+m[2]-1,+m[3],12) : new Date(NaN);
}
function when_(iso) {
  const now = new Date(); let d = now;
  if (iso) {
    const p = parseD_(iso), t = parseD_(fmtD_(now));
    if (isNaN(p)) throw new Error('Tanggal tidak valid.');
    if (p > t) throw new Error('Tanggal tidak boleh di masa depan.');
    if (t - p > 14*864e5) throw new Error('Pengisian mundur maksimal 14 hari.');
    d = p;
  }
  return {date:fmtD_(d), day:HARI[Number(Utilities.formatDate(d,CONFIG.TZ,'u')) % 7],
    time:Utilities.formatDate(now,CONFIG.TZ,'HH:mm:ss'), stamp:now, iso:Utilities.formatDate(d,CONFIG.TZ,'yyyy-MM-dd')};
}

/* ---------- Auth ---------- */
const hash_ = s => Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(b => (b<0?b+256:b).toString(16).padStart(2,'0')).join('');
function newSession_(role,user) {
  const t = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
  PROP().setProperty('SESSION_'+t, JSON.stringify({role:role,user:user,expires:Date.now()+CONFIG.SESSION_HOURS*36e5}));
  return t;
}
function auth_(t, roles) {
  if (!t) throw new Error('Sesi tidak ditemukan. Silakan login kembali.');
  const raw = PROP().getProperty('SESSION_'+t);
  if (!raw) throw new Error('Sesi tidak valid. Silakan login kembali.');
  const s = JSON.parse(raw);
  if (Date.now() > s.expires) { PROP().deleteProperty('SESSION_'+t); throw new Error('Sesi telah kedaluwarsa. Silakan login kembali.'); }
  if (roles.indexOf(s.role) < 0) throw new Error('Anda tidak memiliki akses ke fungsi ini.');
  return s;
}
function cleanup_() {
  const p = PROP(), all = p.getProperties();
  Object.keys(all).forEach(k => { if (k.indexOf('SESSION_') === 0) { try { if (Date.now() > JSON.parse(all[k]).expires) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); } } });
}
function throttle_(key) {
  const c = CacheService.getScriptCache(), n = Number(c.get(key)||0);
  if (n >= 8) throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.');
  return () => c.put(key, String(n+1), 600);
}
function adminHash_() {
  const p = PROP(); let h = p.getProperty('ADMIN_PASSWORD_HASH');
  if (!h) { h = hash_(p.getProperty('ADMIN_PASSWORD') || CONFIG.DEFAULT_ADMIN_PASSWORD); p.setProperty('ADMIN_PASSWORD_HASH',h); p.deleteProperty('ADMIN_PASSWORD'); }
  return h;
}
// Jalankan manual dari editor bila lupa password admin.
function resetAdminPassword() {
  PROP().setProperty('ADMIN_PASSWORD_HASH', hash_(CONFIG.DEFAULT_ADMIN_PASSWORD));
  CacheService.getScriptCache().remove('F_ADMIN');
  return 'Password admin direset ke: ' + CONFIG.DEFAULT_ADMIN_PASSWORD;
}
const pubG_ = g => ({id:g['ID Guru'], nama:g['Nama Guru'], email:g['Email'], username:g['Username']});

function teacherLogin(d) {
  setupDatabase();
  const u = String(d.username||'').trim().toLowerCase(), fail = throttle_('F_T_'+u);
  const g = rows_('guru').find(x => u && (String(x['Username']).trim().toLowerCase() === u || String(x['Email']).trim().toLowerCase() === u));
  if (!g || off_(g) || String(g['PasswordHash']) !== hash_(d.password||'')) { fail(); throw new Error('Username atau password salah, atau akun tidak aktif.'); }
  if (Math.random() < 0.1) cleanup_();
  return {token:newSession_('teacher',g['ID Guru']), teacher:pubG_(g)};
}
function adminLogin(d) {
  setupDatabase();
  const fail = throttle_('F_ADMIN');
  const okU = String(d.username||'').trim().toLowerCase() === String(CONFIG.ADMIN_USER).toLowerCase();
  if (!okU || hash_(String(d.password||'').trim()) !== adminHash_()) { fail(); throw new Error('Username atau password salah.'); }
  cleanup_();
  return {token:newSession_('admin',CONFIG.ADMIN_USER)};
}
function changePassword_(s, d) {
  const np = String(d.newPass||'').trim();
  if (np.length < 8) throw new Error('Password baru minimal 8 karakter.');
  if (s.role === 'admin') {
    if (hash_(String(d.oldPass||'').trim()) !== adminHash_()) throw new Error('Password lama salah.');
    PROP().setProperty('ADMIN_PASSWORD_HASH', hash_(np)); return true;
  }
  const g = rows_('guru').find(x => String(x['ID Guru']) === String(s.user));
  if (!g || String(g['PasswordHash']) !== hash_(d.oldPass||'')) throw new Error('Password lama salah.');
  upsertMany_('guru',[{'ID Guru':s.user,'Password':np}]); return true;
}

/* ---------- Router ---------- */
function doGet() { return out_({ok:true, app:'AKTIVA SMAN 1 NGAWEN API', version:6}); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doPost(e) {
  try {
    const req = JSON.parse((e.postData && e.postData.contents) || '{}');
    return out_({ok:true, status:'success', data:dispatch_(req.action, req.data || {}, req.token || '')});
  } catch (err) { return out_({ok:false, status:'error', message:err && err.message ? err.message : String(err)}); }
}
function dispatch_(a, d, t) {
  switch (a) {
    case 'teacherLogin': return teacherLogin(d);
    case 'adminLogin': return adminLogin(d);
    case 'logout': PROP().deleteProperty('SESSION_'+t); return true;
    case 'changePassword': return changePassword_(auth_(t,['teacher','admin']), d);
    // guru
    case 'teacherHome': return teacherHome_(auth_(t,['teacher']));
    case 'siswaKelas': auth_(t,['teacher']); return rows_('siswa').filter(x => !off_(x) && String(x['ID Kelas']) === String(d.kelasId)).map(x => ({id:x['ID Siswa'], nama:x['Nama Siswa']}));
    case 'teacherExport': return teacherExport_(auth_(t,['teacher']), d);
    case 'sesiDetail': return sesiDetail_(auth_(t,['teacher']), d);
    case 'saveSesi': return saveSesi_(auth_(t,['teacher']), d);
    case 'savePenilaian': return savePenilaian_(auth_(t,['teacher']), d);
    case 'saveRefleksi': return saveRefleksi_(auth_(t,['teacher']), d);
    case 'saveNilai': return saveNilai_(auth_(t,['teacher']), d);
    case 'savePiket': return savePiket_(auth_(t,['teacher']), d);
    case 'saveAgenda': return saveAgenda_(auth_(t,['teacher']), d);
    case 'toggleAgenda': { const s = auth_(t,['teacher']); return agendaRow_(s, d.id, (sh,r,h) => sh.getRange(r, h.indexOf('Selesai')+1).setValue(d.selesai ? 'TRUE' : 'FALSE')); }
    case 'delAgenda': { const s = auth_(t,['teacher']); return agendaRow_(s, d.id, (sh,r) => sh.deleteRow(r)); }
    // admin
    case 'adminDashboard': auth_(t,['admin']); return adminDashboard_();
    case 'adminData': auth_(t,['admin']); return adminData_();
    case 'adminTransaksi': auth_(t,['admin']); return adminTransaksi_(d);
    case 'adminUpsert': auth_(t,['admin']); return adminUpsert_(d);
    case 'adminDelete': auth_(t,['admin']); return adminDelete_(d);
    case 'adminImport': auth_(t,['admin']); return upsertMany_(d.type, d.records || []);
    case 'adminExport': auth_(t,['admin']); return adminExport_(d.type);
  }
  throw new Error('Action tidak dikenal: ' + a);
}

/* ---------- Guru ---------- */
function teacherHome_(s) {
  const n = when_(), mon = n.date.slice(3);
  const g = rows_('guru').find(x => String(x['ID Guru']) === String(s.user));
  if (!g) throw new Error('Sesi tidak valid. Data guru tidak ditemukan.');
  const mine = rows_('jadwal').filter(j => !off_(j) && String(j['ID Guru']) === String(s.user));
  const kelasIds = {}; mine.forEach(j => kelasIds[String(j['ID Kelas'])] = 1);
  const jur = rows_('jurnal').filter(r => String(r['ID Guru']) === String(s.user));
  const pres = rows_('presensi').filter(r => String(r['ID Guru']) === String(s.user) && String(r['Tanggal']).slice(3) === mon);
  const today = {};
  jur.filter(r => r['Tanggal'] === n.date).forEach(r => (today[r['ID Jadwal']] = today[r['ID Jadwal']] || {}).jurnal = true);
  rows_('presensi').filter(r => String(r['ID Guru']) === String(s.user) && r['Tanggal'] === n.date).forEach(r => (today[r['ID Jadwal']] = today[r['ID Jadwal']] || {}).presensi = true);
  const hadir = pres.filter(r => r['Status'] === 'Hadir').length;
  return {
    teacher:pubG_(g), meta:{hari:n.day, tanggal:n.date, iso:n.iso},
    master:{
      mapel:rows_('mapel').filter(x => !off_(x)).map(x => ({id:x['ID Mapel'], nama:x['Mata Pelajaran'], kkm:x['KKM']})),
      kelas:rows_('kelas').filter(x => !off_(x)).map(x => ({id:x['ID Kelas'], nama:x['Kelas']})),
      siswa:rows_('siswa').filter(x => !off_(x) && kelasIds[String(x['ID Kelas'])]).map(x => ({id:x['ID Siswa'], nama:x['Nama Siswa'], kelasId:x['ID Kelas']})),
      jadwal:mine.map(j => ({id:j['ID Jadwal'], mapelId:j['ID Mapel'], kelasId:j['ID Kelas'], hari:j['Hari'], mulai:j['Jam Mulai'], selesai:j['Jam Selesai'], jp:j['Jam Pelajaran']}))
    },
    today:today,
    stats:{jurnal:jur.filter(r => String(r['Tanggal']).slice(3) === mon).length, persenHadir:pres.length ? Math.round(hadir/pres.length*100) : null},
    game:game_(s, jur, mine), agenda:agenda_(s), piket:piket_(s, n),
    refleksiHariIni:rows_('refleksi').some(r => String(r['ID Guru']) === String(s.user) && r['Tanggal'] === n.date),
    recent:jur.slice(-6).reverse().map(r => ({tanggal:r['Tanggal'], mapel:r['Mata Pelajaran'], kelas:r['Kelas'], materi:r['Materi']}))
  };
}
function ownJadwal_(s, id) {
  const j = rows_('jadwal').find(x => String(x['ID Jadwal']) === String(id));
  if (!j || off_(j)) throw new Error('Jadwal tidak ditemukan.');
  if (String(j['ID Guru']) !== String(s.user)) throw new Error('Jadwal ini bukan milik Anda.');
  return j;
}
function resolve_(s, d) {
  if (d.jadwalId) return ownJadwal_(s, d.jadwalId);
  const m = by_(rows_('mapel'),'ID Mapel')[d.mapelId], k = by_(rows_('kelas'),'ID Kelas')[d.kelasId], jp = String(d.jp||'').trim();
  if (!m || off_(m)) throw new Error('Pilih mata pelajaran.');
  if (!k || off_(k)) throw new Error('Pilih kelas.');
  if (!jp) throw new Error('Isi jam pelajaran, contoh: 1-2.');
  return {'ID Jadwal':'M|'+k['ID Kelas']+'|'+m['ID Mapel']+'|'+jp, 'ID Guru':s.user, 'ID Mapel':m['ID Mapel'], 'ID Kelas':k['ID Kelas'], 'Jam Pelajaran':jp};
}
function sesiDetail_(s, d) {
  const j = resolve_(s, d), n = when_(d.tanggal);
  const same = r => r['Tanggal'] === n.date && String(r['ID Jadwal']) === String(j['ID Jadwal']) && String(r['ID Guru']) === String(s.user);
  const jr = rows_('jurnal').find(same), pr = {};
  rows_('presensi').filter(same).forEach(r => pr[r['ID Siswa']] = r['Status']);
  return {materi:jr ? jr['Materi'] : '', keterangan:jr ? jr['Keterangan'] : '', presensi:pr, ada:!!jr};
}
function saveSesi_(s, d) {
  const j = resolve_(s, d), n = when_(d.tanggal);
  const materi = String(d.materi||'').trim();
  if (!materi) throw new Error('Materi pelajaran wajib diisi.');
  const g = by_(rows_('guru'),'ID Guru')[s.user] || {}, m = by_(rows_('mapel'),'ID Mapel')[j['ID Mapel']] || {}, k = by_(rows_('kelas'),'ID Kelas')[j['ID Kelas']] || {};
  const siswa = rows_('siswa').filter(x => !off_(x) && String(x['ID Kelas']) === String(j['ID Kelas']));
  if (!siswa.length) throw new Error('Belum ada siswa aktif di kelas ini. Hubungi admin.');
  const base = {Timestamp:n.stamp, Tanggal:n.date, Hari:n.day, Jam:n.time, 'ID Guru':s.user, 'Nama Guru':g['Nama Guru'], 'ID Mapel':j['ID Mapel'], 'Mata Pelajaran':m['Mata Pelajaran'], 'ID Kelas':j['ID Kelas'], 'Kelas':k['Kelas'], 'ID Jadwal':j['ID Jadwal'], 'Jam Pelajaran':j['Jam Pelajaran']};
  const same = r => r['Tanggal'] === n.date && String(r['ID Jadwal']) === String(j['ID Jadwal']) && String(r['ID Guru']) === String(s.user);
  const st = d.presensi || {}, lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    removeWhere_('jurnal', same); removeWhere_('presensi', same);
    append_('jurnal', [Object.assign({}, base, {Materi:materi, Keterangan:String(d.keterangan||'').trim()})]);
    append_('presensi', siswa.map(x => Object.assign({}, base, {'ID Siswa':x['ID Siswa'], 'Nama Siswa':x['Nama Siswa'], Status:STATUS.indexOf(st[x['ID Siswa']]) >= 0 ? st[x['ID Siswa']] : 'Hadir'})));
  } finally { lock.releaseLock(); }
  return {ok:true, siswa:siswa.length, tanggal:n.date};
}
function savePenilaian_(s, d) {
  const n = when_(), k = by_(rows_('kelas'),'ID Kelas')[d.kelas];
  if (!k) throw new Error('Kelas tidak ditemukan.');
  const v = SCORES.map(x => Number(d[x]));
  if (v.some(x => !(x >= 1 && x <= 5))) throw new Error('Semua nilai harus 1 sampai 5.');
  const total = v.reduce((a,b) => a+b, 0), g = by_(rows_('guru'),'ID Guru')[s.user] || {};
  append_('penilaian', [{Timestamp:n.stamp, Tanggal:n.date, Hari:n.day, Jam:n.time, 'ID Kelas':d.kelas, Kelas:k['Kelas'], Kehadiran:v[0], Penampilan:v[1], Pakaian:v[2], Kebersihan:v[3], Keaktifan:v[4], Perilaku:v[5], Integritas:v[6], 'Total Bintang':total, 'Rata Rata':Number((total/7).toFixed(2)), 'ID Guru':s.user, 'Nama Guru':g['Nama Guru']}]);
  return true;
}
function saveRefleksi_(s, d) {
  const n = when_(), g = by_(rows_('guru'),'ID Guru')[s.user] || {}, r = Number(d.ratingKebahagiaan);
  if (!(r >= 1 && r <= 10)) throw new Error('Rating kebahagiaan harus 1 sampai 10.');
  ['facts','feelings','findings','future'].forEach(f => { if (!String(d[f]||'').trim()) throw new Error('Semua kolom refleksi wajib diisi.'); });
  append_('refleksi', [{Timestamp:n.stamp, Tanggal:n.date, Hari:n.day, Jam:n.time, 'ID Guru':s.user, 'Nama Guru':g['Nama Guru'], Facts:d.facts, Feelings:d.feelings, Findings:d.findings, Future:d.future, 'Rating Kebahagiaan':r}]);
  return true;
}

/* ---------- Admin ---------- */
function adminDashboard_() {
  const n = when_(), mon = n.date.slice(3), act = k => rows_(k).filter(x => !off_(x));
  const G = by_(rows_('guru'),'ID Guru'), M = by_(rows_('mapel'),'ID Mapel'), K = by_(rows_('kelas'),'ID Kelas');
  const jad = act('jadwal').filter(j => j['Hari'] === n.day), done = {};
  rows_('jurnal').filter(r => r['Tanggal'] === n.date).forEach(r => done[String(r['ID Jadwal'])] = 1);
  const belum = jad.filter(j => !done[String(j['ID Jadwal'])]).map(j => ({jp:j['Jam Pelajaran'], mulai:j['Jam Mulai'], guru:(G[j['ID Guru']]||{})['Nama Guru'] || j['ID Guru'], mapel:(M[j['ID Mapel']]||{})['Mata Pelajaran'] || '-', kelas:(K[j['ID Kelas']]||{})['Kelas'] || '-'}));
  const lapor = {}; rows_('laporpiket').filter(r => r['Tanggal'] === n.date).forEach(r => lapor[String(r['ID Guru'])] = 1);
  const piket = act('piket').filter(p => p['Hari'] === n.day).map(p => ({guru:(G[p['ID Guru']]||{})['Nama Guru'] || p['ID Guru'], mulai:p['Jam Mulai'], selesai:p['Jam Selesai'], lapor:!!lapor[String(p['ID Guru'])]}));
  const pres = {}; STATUS.forEach(x => pres[x] = 0);
  rows_('presensi').filter(r => String(r['Tanggal']).slice(3) === mon).forEach(r => { if (pres[r['Status']] !== undefined) pres[r['Status']]++; });
  const rk = {};
  rows_('penilaian').filter(r => String(r['Tanggal']).slice(3) === mon).forEach(r => { const o = rk[r['Kelas']] = rk[r['Kelas']] || {kelas:r['Kelas'], n:0, sum:0}; o.n++; o.sum += Number(r['Rata Rata']||0); });
  return {
    meta:{hari:n.day, tanggal:n.date}, piket:piket,
    counts:{guru:act('guru').length, mapel:act('mapel').length, kelas:act('kelas').length, siswa:act('siswa').length, jadwal:act('jadwal').length},
    hariIni:{jadwal:jad.length, terisi:jad.length - belum.length, belum:belum.sort((a,b) => String(a.mulai).localeCompare(String(b.mulai))).slice(0,60)},
    presensi:pres,
    ranking:Object.keys(rk).map(x => ({kelas:rk[x].kelas, rata:Number((rk[x].sum/rk[x].n).toFixed(2)), n:rk[x].n})).sort((a,b) => b.rata - a.rata).slice(0,8)
  };
}
function adminData_() {
  setupDatabase();
  const o = {};
  Object.keys(MASTER).forEach(k => o[k] = rows_(k).map(r => { const c = Object.assign({}, r); delete c['PasswordHash']; return c; }));
  return o;
}
function adminTransaksi_(d) {
  if (['jurnal','presensi','penilaian','refleksi','nilai','agenda','laporpiket'].indexOf(d.type) < 0) throw new Error('Jenis data tidak dikenal.');
  const q = String(d.q||'').trim().toLowerCase();
  let r = rows_(d.type).slice().reverse();
  if (q) r = r.filter(x => Object.keys(x).some(k => String(x[k]).toLowerCase().indexOf(q) >= 0));
  return {total:r.length, rows:r.slice(0,300)};
}
function adminExport_(type) {
  if (!SHEETS[type]) throw new Error('Export tidak diizinkan.');
  return rows_(type).map(r => { const c = Object.assign({}, r); delete c['PasswordHash']; return c; });
}
function adminUpsert_(d) {
  const r = upsertMany_(d.type, [d.obj || {}]);
  if (r.errors.length) throw new Error(r.errors[0]);
  return r.ids[0];
}
function adminDelete_(d) {
  const c = MASTER[d.type]; if (!c) throw new Error('Jenis master tidak dikenal.');
  const refs = {guru:[['jadwal','ID Guru'],['piket','ID Guru']], mapel:[['jadwal','ID Mapel']], kelas:[['jadwal','ID Kelas'],['siswa','ID Kelas']]}[d.type] || [];
  refs.forEach(p => { if (rows_(p[0]).some(x => String(x[p[1]]) === String(d.id))) throw new Error('Data masih dipakai di ' + SHEETS[p[0]][0] + '. Nonaktifkan (Aktif = FALSE) agar riwayat tetap aman.'); });
  let ok = false; removeWhere_(d.type, x => { if (!ok && String(x[c.id]) === String(d.id)) { ok = true; return true; } return false; });
  return ok;
}
function prep_(type, o, ctx) {
  const t = k => String(o[k] === undefined || o[k] === null ? '' : o[k]).trim();
  if (o['Aktif'] !== undefined) o['Aktif'] = /^(false|0|tidak|nonaktif)$/i.test(t('Aktif')) ? 'FALSE' : 'TRUE';
  if (type === 'guru') {
    if (o['Nama Guru'] !== undefined && !t('Nama Guru')) throw new Error('Nama Guru wajib diisi.');
    if (o['Password']) { if (String(o['Password']).length < 6) throw new Error('Password guru minimal 6 karakter.'); o['PasswordHash'] = hash_(o['Password']); }
    delete o['Password'];
    if (!o['Username'] && o['Email']) o['Username'] = t('Email').split('@')[0];
    if (o['Username'] !== undefined) {
      o['Username'] = t('Username').toLowerCase();
      const same = ctx.users[o['Username']];
      if (o['Username'] && same && String(same) !== String(o['ID Guru'])) throw new Error('Username "' + o['Username'] + '" sudah dipakai.');
    }
  }
  if (type === 'mapel' && o['Mata Pelajaran'] !== undefined && !t('Mata Pelajaran')) throw new Error('Nama mata pelajaran wajib diisi.');
  if (type === 'kelas' && o['Kelas'] !== undefined && !t('Kelas')) throw new Error('Nama kelas wajib diisi.');
  if (type === 'siswa') {
    if (o['Nama Siswa'] !== undefined && !t('Nama Siswa')) throw new Error('Nama siswa wajib diisi.');
    if (o['ID Kelas'] !== undefined) { const k = ctx.kelas[t('ID Kelas')]; if (!k) throw new Error('ID Kelas "' + t('ID Kelas') + '" tidak ditemukan (lihat tab Kelas).'); o['Kelas'] = k['Kelas']; }
  }
  if (type === 'piket') {
    if (o['ID Guru'] !== undefined && !ctx.guru[t('ID Guru')]) throw new Error('ID Guru "' + t('ID Guru') + '" tidak ditemukan.');
    if (o['Hari'] !== undefined) { const h = HARI.find(x => x.toLowerCase() === t('Hari').toLowerCase()); if (!h || h === 'Minggu') throw new Error('Hari harus Senin sampai Sabtu.'); o['Hari'] = h; }
    ['Jam Mulai','Jam Selesai'].forEach(k => { if (o[k] === undefined) return; if (!t(k)) { o[k] = ''; return; } const m = t(k).match(/^(\d{1,2})[:.](\d{2})$/); if (!m || +m[1] > 23 || +m[2] > 59) throw new Error(k + ' harus berformat JJ:MM.'); o[k] = ('0'+m[1]).slice(-2) + ':' + m[2]; });
  }
  if (type === 'jadwal') {
    [['ID Guru','guru'],['ID Mapel','mapel'],['ID Kelas','kelas']].forEach(p => { if (o[p[0]] !== undefined && !ctx[p[1]][t(p[0])]) throw new Error(p[0] + ' "' + t(p[0]) + '" tidak ditemukan.'); });
    if (o['Hari'] !== undefined) { const h = HARI.find(x => x.toLowerCase() === t('Hari').toLowerCase()); if (!h || h === 'Minggu') throw new Error('Hari harus Senin sampai Sabtu.'); o['Hari'] = h; }
    ['Jam Mulai','Jam Selesai'].forEach(k => { if (o[k] !== undefined) { const m = t(k).match(/^(\d{1,2})[:.](\d{2})$/); if (!m || +m[1] > 23 || +m[2] > 59) throw new Error(k + ' harus berformat JJ:MM.'); o[k] = ('0'+m[1]).slice(-2) + ':' + m[2]; } });
  }
}
function upsertMany_(type, recs) {
  const c = MASTER[type]; if (!c) throw new Error('Jenis master tidak dikenal.');
  if (!Array.isArray(recs) || !recs.length) throw new Error('Tidak ada data untuk diproses.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = sheet_(type), h = hdr_(sh), v = sh.getDataRange().getValues(), ic = h.indexOf(c.id), pos = {};
    for (let i = 1; i < v.length; i++) pos[String(v[i][ic]).trim()] = i;
    const ctx = {guru:by_(rows_('guru'),'ID Guru'), mapel:by_(rows_('mapel'),'ID Mapel'), kelas:by_(rows_('kelas'),'ID Kelas'), users:{}};
    rows_('guru').forEach(g => { if (g['Username']) ctx.users[String(g['Username']).toLowerCase()] = g['ID Guru']; });
    const res = {ok:true, inserted:0, updated:0, total:recs.length, errors:[], ids:[]};
    recs.forEach((r, n) => {
      const o = Object.assign({}, r);
      let id = String(o[c.id] || '').trim();
      if (!id) { id = c.p + Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase(); }
      o[c.id] = id;
      try { prep_(type, o, ctx); } catch (e) { res.errors.push('Baris ' + (n+2) + ': ' + e.message); return; }
      const old = pos[id] !== undefined ? v[pos[id]] : null;
      if (!old && o['Aktif'] === undefined) o['Aktif'] = 'TRUE';
      const row = h.map((x,i) => o[x] !== undefined ? o[x] : (old ? old[i] : ''));
      if (old) { v[pos[id]] = row; res.updated++; } else { pos[id] = v.length; v.push(row); res.inserted++; }
      if (type === 'guru') ctx.users[String(o['Username']||'').toLowerCase()] = id;
      if (type === 'kelas') ctx.kelas[id] = {Kelas:o['Kelas']};
      res.ids.push(id);
    });
    if (res.inserted + res.updated) sh.getRange(2,1,v.length-1,h.length).setValues(v.slice(1));
    delete MEMO[type];
    return res;
  } finally { lock.releaseLock(); }
}

/* ---------- Fitur tambahan guru: nilai, agenda, gamifikasi ---------- */
function game_(s, jur, mine) {
  const me = r => String(r['ID Guru']) === String(s.user);
  const pik = rows_('laporpiket').filter(me).length, ref = rows_('refleksi').filter(me).length, pen = rows_('penilaian').filter(me).length, nil = {};
  rows_('nilai').filter(me).forEach(r => nil[r['Tanggal'] + r['ID Kelas'] + r['Jenis'] + r['Topik']] = 1);
  const days = {}; mine.forEach(j => days[j['Hari']] = 1);
  const wk = Object.keys(days).length ? days : {Senin:1, Selasa:1, Rabu:1, Kamis:1, Jumat:1};
  const done = {}; jur.forEach(r => done[r['Tanggal']] = 1);
  let d = parseD_(fmtD_(new Date())), st = 0, first = true;
  for (let i = 0; i < 400; i++, d = new Date(d.getTime() - 864e5)) {
    if (!wk[HARI[d.getDay()]]) continue;
    if (done[fmtD_(d)]) st++; else if (!first) break;
    first = false;
  }
  const nn = Object.keys(nil).length;
  return {poin:jur.length*10 + ref*15 + pen*5 + nn*10 + pik*10, streak:st, jurnal:jur.length, refleksi:ref, nilai:nn};
}
function agenda_(s) {
  const a = rows_('agenda').filter(r => String(r['ID Guru']) === String(s.user)).sort((x, y) => String(x['Tanggal']).localeCompare(String(y['Tanggal'])))
    .map(r => ({id:r['ID Agenda'], tanggal:r['Tanggal'], judul:r['Judul'], jenis:r['Jenis'], selesai:String(r['Selesai']).toUpperCase() === 'TRUE'}));
  return a.filter(x => !x.selesai).concat(a.filter(x => x.selesai).slice(-5));
}
function saveAgenda_(s, d) {
  const judul = String(d.judul||'').trim(), tgl = String(d.tanggal||'').trim();
  if (!judul) throw new Error('Judul agenda wajib diisi.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tgl)) throw new Error('Tanggal agenda tidak valid.');
  append_('agenda', [{Timestamp:new Date(), 'ID Agenda':'A' + Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase(), 'ID Guru':s.user, Tanggal:tgl, Judul:judul.slice(0,200), Jenis:String(d.jenis||'Pengingat'), Selesai:'FALSE'}]);
  return true;
}
function agendaRow_(s, id, fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = sheet_('agenda'), v = sh.getDataRange().getValues(), h = v[0].map(String), ii = h.indexOf('ID Agenda'), ig = h.indexOf('ID Guru');
    for (let i = 1; i < v.length; i++) if (String(v[i][ii]) === String(id) && String(v[i][ig]) === String(s.user)) { fn(sh, i+1, h); delete MEMO.agenda; return true; }
  } finally { lock.releaseLock(); }
  throw new Error('Agenda tidak ditemukan.');
}
function saveNilai_(s, d) {
  const n = when_(), k = by_(rows_('kelas'),'ID Kelas')[d.kelasId], m = by_(rows_('mapel'),'ID Mapel')[d.mapelId];
  if (!k || !m) throw new Error('Pilih kelas dan mata pelajaran.');
  const jenis = String(d.jenis||'').trim(), topik = String(d.topik||'').trim();
  if (!jenis || !topik) throw new Error('Jenis dan topik penilaian wajib diisi.');
  const sc = d.nilai || {}, kkm = Number(m['KKM']) || 75, g = by_(rows_('guru'),'ID Guru')[s.user] || {};
  const siswa = rows_('siswa').filter(x => !off_(x) && String(x['ID Kelas']) === String(k['ID Kelas']) && sc[x['ID Siswa']] !== undefined && sc[x['ID Siswa']] !== '');
  if (!siswa.length) throw new Error('Isi minimal satu nilai siswa.');
  siswa.forEach(x => { const v = Number(sc[x['ID Siswa']]); if (!(v >= 0 && v <= 100)) throw new Error('Nilai ' + x['Nama Siswa'] + ' harus 0 sampai 100.'); });
  const same = r => r['Tanggal'] === n.date && String(r['ID Guru']) === String(s.user) && String(r['ID Kelas']) === String(k['ID Kelas']) && String(r['ID Mapel']) === String(m['ID Mapel']) && r['Jenis'] === jenis && r['Topik'] === topik;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    removeWhere_('nilai', same);
    append_('nilai', siswa.map(x => ({Timestamp:n.stamp, Tanggal:n.date, 'ID Guru':s.user, 'Nama Guru':g['Nama Guru'], 'ID Mapel':m['ID Mapel'], 'Mata Pelajaran':m['Mata Pelajaran'], 'ID Kelas':k['ID Kelas'], Kelas:k['Kelas'], Jenis:jenis, Topik:topik, 'ID Siswa':x['ID Siswa'], 'Nama Siswa':x['Nama Siswa'], Nilai:Number(sc[x['ID Siswa']]), KKM:kkm, Tuntas:Number(sc[x['ID Siswa']]) >= kkm ? 'Ya' : 'Remedial'})));
  } finally { lock.releaseLock(); }
  const v = siswa.map(x => Number(sc[x['ID Siswa']]));
  return {siswa:v.length, tuntas:v.filter(x => x >= kkm).length, rata:Number((v.reduce((a,b) => a+b, 0)/v.length).toFixed(1))};
}

/* ---------- Unduh laporan guru (PDF / Excel dibuat di browser) ---------- */
const EXPORT_COLS = {
  jurnal:['Tanggal','Hari','Mata Pelajaran','Kelas','Jam Pelajaran','Materi','Keterangan'],
  presensi:['Tanggal','Mata Pelajaran','Kelas','Jam Pelajaran','Nama Siswa','Status'],
  nilai:['Tanggal','Mata Pelajaran','Kelas','Jenis','Topik','Nama Siswa','Nilai','KKM','Tuntas'],
  penilaian:['Tanggal','Kelas','Kehadiran','Penampilan','Pakaian','Kebersihan','Keaktifan','Perilaku','Integritas','Rata Rata'],
  refleksi:['Tanggal','Facts','Feelings','Findings','Future','Rating Kebahagiaan'],
  agenda:['Tanggal','Judul','Jenis','Selesai']
};
function teacherExport_(s, d) {
  const cols = EXPORT_COLS[d.type];
  if (!cols) throw new Error('Jenis laporan tidak dikenal.');
  const iso = v => { const x = parseD_(v); return isNaN(x) ? '' : Utilities.formatDate(x, CONFIG.TZ, 'yyyy-MM-dd'); };
  const yes = v => v === true || String(v).toUpperCase() === 'TRUE' ? 'Ya' : (v === false || String(v).toUpperCase() === 'FALSE' ? 'Tidak' : v);
  let r = rows_(d.type).filter(x => String(x['ID Guru']) === String(s.user));
  if (/^\d{4}-\d{2}-\d{2}$/.test(d.dari || '')) r = r.filter(x => iso(x['Tanggal']) >= d.dari);
  if (/^\d{4}-\d{2}-\d{2}$/.test(d.sampai || '')) r = r.filter(x => iso(x['Tanggal']) <= d.sampai);
  if (d.kelasId && cols.indexOf('Kelas') >= 0) r = r.filter(x => String(x['ID Kelas']) === String(d.kelasId));
  if (d.mapelId && cols.indexOf('Mata Pelajaran') >= 0) r = r.filter(x => String(x['ID Mapel']) === String(d.mapelId));
  r.sort((a, b) => (iso(a['Tanggal']) + String(a['Timestamp'] || '')).localeCompare(iso(b['Tanggal']) + String(b['Timestamp'] || '')));
  const total = r.length, LIM = 5000;
  const g = rows_('guru').find(x => String(x['ID Guru']) === String(s.user)) || {};
  return {cols: cols, total: total, terpotong: total > LIM, guru: g['Nama Guru'] || '',
    rows: r.slice(0, LIM).map(x => cols.map(c => { const v = (c === 'Selesai' || c === 'Tuntas') ? yes(x[c]) : x[c]; return v === undefined || v === null ? '' : v; }))};
}

/* ---------- Piket guru ---------- */
function piket_(s, n) {
  const G = by_(rows_('guru'),'ID Guru'), all = rows_('piket').filter(p => !off_(p)), me = x => String(x['ID Guru']) === String(s.user);
  const today = all.filter(p => p['Hari'] === n.day);
  const lp = rows_('laporpiket').find(r => r['Tanggal'] === n.date && me(r));
  return {
    bertugas: today.map(p => ({guru:(G[p['ID Guru']]||{})['Nama Guru'] || p['ID Guru'], mulai:p['Jam Mulai'], selesai:p['Jam Selesai'], ket:p['Keterangan'], saya:me(p)}))
      .sort((a,b) => String(a.mulai).localeCompare(String(b.mulai))),
    saya: today.some(me),
    hariSaya: all.filter(me).map(p => p['Hari']),
    laporan: lp ? {terlambat:lp['Siswa Terlambat'], izin:lp['Siswa Izin Sakit'], guruAbsen:lp['Guru Tidak Hadir'], kejadian:lp['Kejadian'], tindak:lp['Tindak Lanjut']} : null
  };
}
function savePiket_(s, d) {
  const n = when_();
  const jadwal = rows_('piket').some(p => !off_(p) && p['Hari'] === n.day && String(p['ID Guru']) === String(s.user));
  if (!jadwal) throw new Error('Anda tidak terjadwal piket hari ini. Hubungi admin jika menggantikan rekan.');
  const g = by_(rows_('guru'),'ID Guru')[s.user] || {}, f = k => String(d[k]||'').trim().slice(0,1000) || 'Nihil';
  const same = r => r['Tanggal'] === n.date && String(r['ID Guru']) === String(s.user);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    removeWhere_('laporpiket', same);
    append_('laporpiket', [{Timestamp:n.stamp, Tanggal:n.date, Hari:n.day, Jam:n.time, 'ID Guru':s.user, 'Nama Guru':g['Nama Guru'],
      'Siswa Terlambat':f('terlambat'), 'Siswa Izin Sakit':f('izin'), 'Guru Tidak Hadir':f('guruAbsen'), Kejadian:f('kejadian'), 'Tindak Lanjut':f('tindak')}]);
  } finally { lock.releaseLock(); }
  return true;
}
