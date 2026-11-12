/******************************************************
 * ระบบนิเทศและประเมินการสอน — โรงเรียนในสังกัด อบจ.ศรีสะเกษ
 * Code.gs  (Google Apps Script)
 * ฐานข้อมูล: Google Sheets (ชีตละ collection: id | json | updated)
 * ไฟล์/ภาพ/PDF: Google Drive   |   Firestore: ตัวเลือกเสริม (ตั้งค่าในหน้าเว็บ)
 ******************************************************/
const APP_NAME = 'ระบบนิเทศและประเมินการสอน';
const ROOT_FOLDER_ID = '';           // ใส่ ID โฟลเดอร์ Drive หลัก (ว่าง = สร้างให้อัตโนมัติ)
const COLLS = ['Users','Evaluations','LessonPlans','Bookings','Settings','Schools'];
const SESSION_HOURS = 5;

/* ---------------- Web entry ---------------- */
function doGet(e){
  ensureSetup_();
  const t = HtmlService.createTemplateFromFile('index');
  const boot = { settings: getSettings_(), url: ScriptApp.getService().getUrl(), portal: String((e&&e.parameter&&e.parameter.p)||'').replace(/[^a-z]/g,''), original: isOriginal_() };
  const sk = String((e&&e.parameter&&e.parameter.s)||'').replace(/[^A-Za-z0-9]/g,'');   /* ลิงก์ประเมินความพึงพอใจ (ไม่ต้องล็อกอิน) */
  if (sk){ boot.survey = sk; boot.schools = Object.values(readColl_('Schools')).map(x=>x.name).filter(Boolean); }
  try{ boot.logo_src = String(((boot.settings||{}).school_info||{}).logo||''); boot.logo_data = logoData_(boot.logo_src); }catch(err){}
  t.boot = JSON.stringify(boot).replace(/</g,'\\u003c');
  return t.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- Sheet helpers ---------------- */
function ss_(){ return SpreadsheetApp.getActive(); }
function sh_(name){
  const s = ss_().getSheetByName(name);
  if (s) return s;
  const n = ss_().insertSheet(name);
  n.getRange(1,1,1,3).setValues([['id','json','updated']]).setFontWeight('bold');
  n.setFrozenRows(1);
  return n;
}
function readColl_(name){
  const v = sh_(name).getDataRange().getValues(), out = {};
  for (let i=1;i<v.length;i++){ if(!v[i][0]) continue;
    try{ out[v[i][0]] = JSON.parse(v[i][1]||'{}'); }catch(e){} }
  return out;
}
/* ops: [{op:'set'|'merge'|'delete', id, data}] — เขียนทีเดียวทั้งชีต */
function writeColl_(name, ops){
  const s = sh_(name), v = s.getDataRange().getValues(), idx = {};
  for (let i=1;i<v.length;i++) if (v[i][0]) idx[v[i][0]] = i;
  const now = new Date();
  ops.forEach(o=>{
    const i = idx[o.id];
    if (o.op==='delete'){ if(i!=null) v[i] = ['','','']; return; }
    const old = (i!=null && o.op==='merge') ? JSON.parse(v[i][1]||'{}') : {};
    const json = JSON.stringify(Object.assign(old, o.data||{}));
    if (json.length > 49000) throw new Error('ข้อมูลใหญ่เกิน 50,000 ตัวอักษร/เซลล์');
    if (i!=null) v[i] = [o.id, json, now]; else { idx[o.id] = v.length; v.push([o.id, json, now]); }
  });
  const rows = v.slice(1).filter(r=>r[0]);
  s.getRange(2,1,Math.max(s.getLastRow(),2),3).clearContent();
  if (rows.length) s.getRange(2,1,rows.length,3).setValues(rows);
}
function uid_(p){ return (p||'') + Date.now().toString(36) + Math.random().toString(36).slice(2,6); }
function hash_(s){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s))
    .map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
}

/* ---------------- Setup / default data ---------------- */
/* ข้อมูลเริ่มต้น: ศน. และโรงเรียนในความรับผิดชอบ (สร้างครั้งเดียว ถ้ายังไม่มี) */
const DEFAULT_SCHOOLS = [
  {id:'S01', code:'S01', name:'โรงเรียนไพรบึงวิทยาคม', size:'ขนาดใหญ่'},
  {id:'S02', code:'S02', name:'โรงเรียนโนนปูนวิทยาคม', size:'ขนาดเล็ก'},
  {id:'S03', code:'S03', name:'โรงเรียนกุญชรศิรวิทย์', size:'ขนาดเล็ก'},
  {id:'S04', code:'S04', name:'โรงเรียนหนองถ่มวิทยา', size:'ขนาดเล็ก'}];
const DEFAULT_SUPERVISOR = { username:'piyaphan', name:'นางปิยะพันธ์ พิชญ์ประเสริฐ', position:'ศึกษานิเทศก์',
  rank:'ศึกษานิเทศก์เชี่ยวชาญ', role:'Supervisor', status:'active', school:'' };
function seedDefaults_(){
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SEED_V1') || !isOriginal_()) return;
  const sc = readColl_('Schools'), names = Object.keys(sc).map(k=>sc[k].name), sOps = [];
  DEFAULT_SCHOOLS.forEach(x=>{ if (names.indexOf(x.name)<0 && !sc[x.id]) sOps.push({op:'set', id:x.id, data:Object.assign({district:'', director:'', note:''}, x)}); });
  if (sOps.length) writeColl_('Schools', sOps);
  const us = readColl_('Users');
  if (!us[DEFAULT_SUPERVISOR.username]){
    writeColl_('Users', [{op:'set', id:DEFAULT_SUPERVISOR.username, data:Object.assign({}, DEFAULT_SUPERVISOR,
      { pass:hash_(DEFAULT_SUPERVISOR.username), schools:DEFAULT_SCHOOLS.map(x=>x.name), group:'', signature_url:'' })}]);
  }
  p.setProperty('SEED_V1','1');
}
function seedV2_(){   /* เปลี่ยนชื่อหน่วยงานหลักเป็นกลุ่มงานนิเทศการศึกษา */
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SEED_V2') || !isOriginal_()) return;
  const st = readColl_('Settings').main;
  if (st){ const si = st.school_info || {};
    if (!si.name || si.name==='โรงเรียนกระดุมทองวิทยา'){
      si.name='กลุ่มงานนิเทศการศึกษา'; si.affiliation='กองการศึกษา ศาสนา และวัฒนธรรม องค์การบริหารส่วนจังหวัดศรีสะเกษ';
      writeColl_('Settings',[{op:'merge', id:'main', data:{school_info:si}}]); } }
  p.setProperty('SEED_V2','1');
}
function ensureSetup_(){
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SETUP_DONE')) {
    const us = readColl_('Users');   /* ไฟล์สำเนา template: ถ้าไม่มีแอดมิน ให้สร้างใหม่ */
    if (!Object.keys(us).some(function(k){ return us[k].role==='Admin'; })) p.deleteProperty('SETUP_DONE');
    else { seedDefaults_(); seedV2_(); migrateFoldersV2_(); return; }
  }
  COLLS.forEach(sh_);
  const users = readColl_('Users');
  if (!Object.keys(users).some(function(k){ return users[k].role==='Admin'; })){
    writeColl_('Users', [{op:'set', id:'admin', data:{
      username:'admin', pass:hash_('admin1234'), name:'ผู้ดูแลระบบ', role:'Admin',
      status:'active', school:'', group:'', signature_url:''}}]);
  }
  const st = readColl_('Settings');
  if (!st.main) writeColl_('Settings', [{op:'set', id:'main', data: defaultSettings_()}]);
  p.setProperty('SETUP_DONE','1');
  seedDefaults_(); seedV2_(); migrateFoldersV2_();
}
function defaultSettings_(){
  return {
    academic_year: '2569', term: '2',
    developer_name: 'นายนนทพัทธ์ หิรัญเรือง',
    school_info: { name:'กลุ่มงานนิเทศการศึกษา', affiliation:'กองการศึกษา ศาสนา และวัฒนธรรม องค์การบริหารส่วนจังหวัดศรีสะเกษ',
                   logo:'', director:'', schools:['โรงเรียนกระดุมทองวิทยา'] },
    periods: [
      {no:1,time:'08:30-09:20'},{no:2,time:'09:20-10:10'},{no:3,time:'10:10-11:00'},{no:4,time:'11:00-11:50'},
      {no:5,time:'12:50-13:40'},{no:6,time:'13:40-14:30'},{no:7,time:'14:30-15:20'},{no:8,time:'15:20-16:10'}],
    locked_dates: [],
    theme: { primary:'#F97316', primary2:'#EA580C', accent:'#FACC15' },
    grade_scale: [
      {min:90,label:'ดีเยี่ยม'},{min:80,label:'ดีมาก'},{min:70,label:'ดี'},{min:60,label:'พอใช้'},{min:0,label:'ปรับปรุง'}],
    rubrics: null      // null = ใช้ค่าเริ่มต้นในหน้าเว็บ
  };
}
function getSettings_(){
  const s = readColl_('Settings').main || defaultSettings_();
  return Object.assign(defaultSettings_(), s);
}

/* ---------------- Auth (session token) ---------------- */
const PORTAL_ROLES = { teacher:['Teacher','User'], supervisor:['Supervisor'], admin:['Admin','SchoolAdmin'] };
const PORTAL_TH = { teacher:'ครูผู้สอน', supervisor:'ศึกษานิเทศก์', admin:'แอดมินระบบ / แอดมินโรงเรียน' };
function apiLogin(username, password, portal){
  ensureSetup_();
  const users = readColl_('Users');
  const un = String(username).trim().toLowerCase();
  const u = Object.values(users).find(x=>String(x.username).toLowerCase()===un);
  if (!u || u.pass !== hash_(password)) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  if (u.status !== 'active') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
  if (portal && PORTAL_ROLES[portal] && PORTAL_ROLES[portal].indexOf(u.role)<0){
    const right = Object.keys(PORTAL_ROLES).find(k=>PORTAL_ROLES[k].indexOf(u.role)>=0);
    throw new Error('บัญชีนี้ไม่ได้อยู่ในส่วน "'+PORTAL_TH[portal]+'" — กรุณาเข้าสู่ระบบที่การ์ด "'+PORTAL_TH[right]+'"');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('S_'+token, u.username, SESSION_HOURS*3600);
  return JSON.stringify({ ok:true, token, user: publicUser_(u), data: loadFor_(u), warn: u.role==='Admin' ? deployWarn_() : null });
}
function apiResume(token){
  const u = me_(token);
  return JSON.stringify({ ok:true, user: publicUser_(u), data: loadFor_(u), warn: u.role==='Admin' ? deployWarn_() : null });
}
function apiLogout(token){ CacheService.getScriptCache().remove('S_'+token); return '{"ok":true}'; }
function me_(token){
  const name = CacheService.getScriptCache().get('S_'+token);
  if (!name) throw new Error('AUTH_EXPIRED: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const u = Object.values(readColl_('Users')).find(x=>x.username===name);
  if (!u || u.status!=='active') throw new Error('AUTH_EXPIRED: บัญชีไม่พร้อมใช้งาน');
  return u;
}
/* เปลี่ยนชื่อผู้ใช้ (เฉพาะ Admin): ย้ายบัญชี แผน ผลประเมิน การจอง ไปใช้ชื่อใหม่ รหัสผ่านเดิมคงไว้ */
function apiRenameUser(token, oldUn, newUn){
  const u = me_(token); need_(u,['Admin']);
  oldUn = String(oldUn||''); newUn = String(newUn||'').replace(/[\s\u00A0\u200B-\u200F\u2028-\u202F\u2060\uFEFF]/g,'');
  if (!/^[\p{L}\p{M}\p{N}._@\-]{1,40}$/u.test(newUn)) throw new Error('ชื่อผู้ใช้ใหม่ไม่ถูกต้อง');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const users = readColl_('Users'), old = users[oldUn];
    if (!old) throw new Error('ไม่พบผู้ใช้เดิม');
    if (newUn===oldUn) throw new Error('ชื่อใหม่ซ้ำกับชื่อเดิม');
    if (Object.keys(users).some(function(k){ return k!==oldUn && k.toLowerCase()===newUn.toLowerCase(); })) throw new Error('ชื่อผู้ใช้นี้มีแล้ว');
    const uops = [{op:'delete', id:oldUn}, {op:'set', id:newUn, data:Object.assign({}, old, {username:newUn, updated_by:u.username, updated_at:new Date().toISOString()})}];
    Object.keys(users).forEach(function(k){ if (users[k].linked_user===oldUn) uops.push({op:'merge', id:k, data:{linked_user:newUn}}); });
    const cnt = {plans:0, evals:0, bookings:0};
    const lp = readColl_('LessonPlans'), lops = [];
    Object.keys(lp).forEach(function(k){ if (lp[k].teacher_id===oldUn){ lops.push({op:'merge', id:k, data:{teacher_id:newUn}}); cnt.plans++; } });
    const ev = readColl_('Evaluations'), eops = [];
    Object.keys(ev).forEach(function(k){ const x = ev[k], d = {};
      if (x.teacher_id===oldUn) d.teacher_id = newUn;
      if (x.evaluator1===oldUn) d.evaluator1 = newUn;
      if (x.evaluator2===oldUn) d.evaluator2 = newUn;
      if (x.updated_by===oldUn) d.updated_by = newUn;
      if ((x.evaluators||[]).indexOf(oldUn)>=0) d.evaluators = x.evaluators.map(function(e){ return e===oldUn?newUn:e; });
      if (Object.keys(d).length){ eops.push({op:'merge', id:k, data:d}); cnt.evals++; } });
    const bk = readColl_('Bookings'), bops = [];
    Object.keys(bk).forEach(function(k){ if (bk[k].teacher_id===oldUn){ bops.push({op:'merge', id:k, data:{teacher_id:newUn}}); cnt.bookings++; } });
    writeColl_('Users', uops);
    if (lops.length) writeColl_('LessonPlans', lops);
    if (eops.length) writeColl_('Evaluations', eops);
    if (bops.length) writeColl_('Bookings', bops);
    return JSON.stringify({ok:true, plans:cnt.plans, evals:cnt.evals, bookings:cnt.bookings});
  } finally { lock.releaseLock(); }
}
function need_(u, roles){ if (roles.indexOf(u.role)<0) throw new Error('ไม่มีสิทธิ์ทำรายการนี้'); }
function publicUser_(u){ const c = Object.assign({}, u); delete c.pass; return c; }

/* โหลดข้อมูลตามบทบาท (RBAC ฝั่งเซิร์ฟเวอร์)
   Admin = ทุกอย่าง | Supervisor = เฉพาะโรงเรียนที่รับผิดชอบ (ว่าง = ทุกโรงเรียน)
   User (ผู้บริหาร) = ทั้งโรงเรียนของตน | Teacher = ของตนเอง */
function scope_(u){
  if (u.role==='Admin') return null;
  if (u.role==='Supervisor') return (u.schools && u.schools.length) ? u.schools : null;
  return [u.school||''];   /* Teacher / User / SchoolAdmin */
}
function evBy_(x, un){ return (x.evaluators||[]).indexOf(un)>=0 || x.evaluator1===un || x.evaluator2===un || x.updated_by===un; }
function loadFor_(u){
  const users = readColl_('Users'), ev = readColl_('Evaluations'),
        lp = readColl_('LessonPlans'), bk = readColl_('Bookings'), sc = readColl_('Schools');
  const sc_ = scope_(u), inS = s => !sc_ || sc_.indexOf(s||'')>=0;
  const pubUsers = {};
  Object.keys(users).forEach(k=>{ const x = publicUser_(users[k]);
    const staff = x.role==='Admin' || x.role==='Supervisor';
    if (u.role!=='Admin' && !staff && !inS(x.school)) return;
    if (u.role!=='Admin'){ delete x.status; }
    if (u.role==='Supervisor' && x.role==='Supervisor' && k!==u.username) return;   /* ศน. ไม่เห็น ศน. คนอื่น */
    pubUsers[k]=x; });
  const schoolOf = id => (users[id]||{}).school || '';
  const pick = (obj, f) => { const o = {}; Object.keys(obj).forEach(k=>{ if (f(obj[k])) o[k]=obj[k]; }); return o; };
  let E, L, B, S;
  if (u.role==='Teacher'){
    E = pick(ev, x=>x.teacher_id===u.username);
    L = pick(lp, x=>x.teacher_id===u.username);
  } else if (u.role==='User'){   /* ผอ./รองผอ.: ดูภาพรวมทั้งโรงเรียนของตน (อ่านอย่างเดียว) */
    E = pick(ev, x=>inS(x.school || schoolOf(x.teacher_id)));
    L = pick(lp, x=>inS(schoolOf(x.teacher_id)));
  } else {
    E = pick(ev, x=>inS(x.school || schoolOf(x.teacher_id)) && (u.role!=='Supervisor' || evBy_(x, u.username)));
    L = pick(lp, x=>inS(schoolOf(x.teacher_id)) && (u.role!=='Supervisor' || x.status!=='ประเมินแล้ว' || !x.eval_id || !!E[x.eval_id]));
  }
  B = pick(bk, x=>inS(x.school || schoolOf(x.teacher_id)));
  if (u.role==='Teacher'){   /* ครูไม่เห็นชื่อผู้จองคนอื่น */
    Object.keys(B).forEach(k=>{ if (B[k].teacher_id!==u.username) B[k] = Object.assign({}, B[k], {teacher_id:'(จองแล้ว)', note:''}); });
    Object.keys(pubUsers).forEach(k=>{ const x=pubUsers[k]; if (k!==u.username && x.role!=='Supervisor' && x.role!=='Admin') delete pubUsers[k]; });
  }
  S = pick(sc, x=>inS(x.name));
  const SV = (u.role==='Admin'||u.role==='Supervisor') ? pick(readColl_('Surveys'), x=>!x.school||inS(x.school)) : {};
  return { Users:pubUsers, Evaluations:E, LessonPlans:L, Bookings:B, Schools:S, Surveys:SV, settings:getSettings_() };
}

/* ---------------- Generic commit with RBAC ---------------- */
// ops: [{c, op, id, data}]
function apiCommit(token, opsJson){
  const u = me_(token);
  const ops = JSON.parse(opsJson);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const by = {};
    ops.forEach(o=>{
      authorize_(u, o);
      if (o.op!=='delete' && o.data){ o.data.updated_by = u.username; o.data.updated_at = new Date().toISOString(); }
      if (o.c==='Users' && o.data && o.data.newPass){ o.data.pass = hash_(o.data.newPass); delete o.data.newPass; }
      (by[o.c] = by[o.c]||[]).push(o);
    });
    const lpPrev = by.LessonPlans ? readColl_('LessonPlans') : null;
    const stPrev = (by.Settings && u.role==='Admin') ? (readColl_('Settings').main || {}) : null;
    Object.keys(by).forEach(c=>writeColl_(c, by[c]));
    if (lpPrev) trashReplacedPlanFiles_(u, lpPrev, by.LessonPlans);
    if (stPrev) purgeRemovedTeacherBlocks_(u, stPrev, by.Settings);
    return JSON.stringify({ ok:true });
  } finally { lock.releaseLock(); }
}
function authorize_(u, o){
  if (COLLS.indexOf(o.c)<0) throw new Error('ไม่รู้จักชุดข้อมูล '+o.c);
  const r = u.role;
  if (o.c==='Users'){
    if (r==='Admin') return;
    if (r==='SchoolAdmin' && o.id!==u.username){   /* แอดมินโรงเรียน: จัดการครู/ผู้บริหารในโรงเรียนตนเอง */
      if (o.op==='delete') throw new Error('แอดมินโรงเรียนลบผู้ใช้ไม่ได้ (ใช้ระงับแทน)');
      const cur = readColl_('Users')[o.id], ok = ['Teacher','User'];
      if (cur && (cur.school!==u.school || ok.indexOf(cur.role)<0)) throw new Error('แก้ไขได้เฉพาะครู/ผู้บริหารในโรงเรียนของท่าน');
      if (o.data){ if (o.data.role && ok.indexOf(o.data.role)<0) throw new Error('กำหนดบทบาทได้เฉพาะครูหรือผู้บริหาร');
        if (!cur && !o.data.role) o.data.role='Teacher';
        o.data.school = u.school; delete o.data.schools; delete o.data.pass; }
      return;
    }
    if (o.id===u.username && o.op==='merge'){ // แก้โปรไฟล์/ลายเซ็นตนเอง
      ['role','status','username','pass'].forEach(k=>{ if (o.data) delete o.data[k]; }); return; }
    throw new Error('ไม่มีสิทธิ์จัดการผู้ใช้');
  }
  if (o.c==='Settings' && r==='SchoolAdmin' && o.op==='merge' && o.data && Object.keys(o.data).join()==='locked_dates'){
    const others = a => JSON.stringify((a||[]).filter(d=>d.school!==u.school));
    if (others(getSettings_().locked_dates)!==others(o.data.locked_dates)) throw new Error('ล็อกวันได้เฉพาะโรงเรียนของท่าน');
    return;
  }
  if (o.c==='Schools' && r==='SchoolAdmin' && o.op==='merge' && o.data && Object.keys(o.data).join()==='periods'){
    const cur = readColl_('Schools')[o.id]; if (cur && cur.name===u.school) return;
    throw new Error('ตั้งค่าคาบเรียนได้เฉพาะโรงเรียนของท่าน'); }
  if (o.c==='Settings' && r==='Supervisor' && o.op==='merge' && o.data && ['cert','survey','rubrics'].indexOf(Object.keys(o.data).join())>=0) return;   /* ศน. ตั้งค่าเกียรติบัตร/แบบประเมินความพึงพอใจ/ข้อกระทงการนิเทศได้ */
  if (o.c==='Settings' || o.c==='Schools'){ need_(u,['Admin']); return; }
  if (o.c==='Evaluations'){ need_(u,['Admin','Supervisor']);
    if (r==='Supervisor'){ const cur = readColl_('Evaluations')[o.id];
      if (cur && !evBy_(cur, u.username)) throw new Error('แก้ไขได้เฉพาะผลการนิเทศของท่าน');
      if (!cur && o.data && o.data.evaluators && o.data.evaluators.indexOf(u.username)<0) throw new Error('ต้องมีชื่อท่านเป็นผู้นิเทศ'); }
    const sc = scope_(u); if (sc && o.data && o.data.school && sc.indexOf(o.data.school)<0) throw new Error('โรงเรียนนี้ไม่อยู่ในความรับผิดชอบของท่าน');
    return; }
  if (o.c==='LessonPlans'){
    if (r==='Admin'||r==='Supervisor') return;
    if (o.op==='delete') { const cur = readColl_('LessonPlans')[o.id]; if (cur && cur.teacher_id===u.username && cur.status==='รอประเมิน') return; }
    if (o.data && o.data.teacher_id===u.username) return;
    throw new Error('ส่งแผนได้เฉพาะของตนเอง');
  }
  if (o.c==='Bookings'){
    if (r==='Admin'||r==='Supervisor') return;
    if (r==='SchoolAdmin'){
      const cur = readColl_('Bookings')[o.id];
      if (o.op==='delete' ? (cur && cur.school===u.school) : (o.data && o.data.school===u.school)) return;
      throw new Error('จัดการได้เฉพาะโรงเรียนของท่าน');
    }
    if (o.op==='delete'){ const cur = readColl_('Bookings')[o.id]; if (cur && cur.teacher_id===u.username) return; throw new Error('ยกเลิกได้เฉพาะการจองของตนเอง'); }
    if (o.data && o.data.teacher_id===u.username && !o.data.is_locked){
      o.data.school = u.school || '';
      const all = readColl_('Bookings');
      const dup = Object.keys(all).some(k=>k!==o.id && all[k].date===o.data.date && +all[k].period_no===+o.data.period_no && (all[k].school||'')===(o.data.school||''));
      if (dup) throw new Error('คาบนี้ถูกจองแล้ว');
      const st = getSettings_();
      if ((st.locked_dates||[]).some(d=>d.date===o.data.date && (!d.school || d.school===o.data.school))) throw new Error('วันนี้โรงเรียนงดนิเทศ');
      return;
    }
    throw new Error('ไม่มีสิทธิ์จอง/ล็อกคาบนี้');
  }
}

/* ---------------- Import ข้อมูลหลัก (.xlsx แปลงเป็น JSON ที่หน้าเว็บ) ----------------
   payload: { schools:[{name,code,size,district,director,note}],
              users:[{username,password,name,position,rank,group,school,schools[],role,status}],
              replaceSchools:false } */
function apiImport(token, json){
  const u = me_(token); need_(u,['Admin']);
  const p = JSON.parse(json);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const now = new Date().toISOString();
    const curS = readColl_('Schools'), byName = {};
    Object.keys(curS).forEach(k=>byName[curS[k].name]=k);
    const sOps = [];
    if (p.replaceSchools) Object.keys(curS).forEach(k=>sOps.push({op:'delete', id:k}));
    let sNew=0, sUpd=0;
    (p.schools||[]).forEach(x=>{
      if (!x.name) return;
      const id = (!p.replaceSchools && byName[x.name]) || x.code || ('SC'+Utilities.getUuid().slice(0,8));
      if (!p.replaceSchools && byName[x.name]) sUpd++; else sNew++;
      sOps.push({op:'merge', id:id, data:{id:id, code:x.code||'', name:x.name, size:x.size||'', district:x.district||'',
        director:x.director||'', note:x.note||'', updated_at:now, updated_by:u.username}});
    });
    if (sOps.length) writeColl_('Schools', sOps);
    const curU = readColl_('Users'), uOps = [];
    let uNew=0, uUpd=0, skip=[];
    const nn = function(s){ return String(s||'').trim().replace(/^(นางสาว|น\.ส\.|นาง|นาย|ดร\.|ว่าที่\s*ร\.ต\.|ว่าที่ร้อยตรี)\s*/,'').replace(/[\s.]/g,''); };
    const byKey = {}; Object.keys(curU).forEach(function(k){ const x = curU[k];
      if (x.role==='Supervisor') byKey['S|'+nn(x.name)] = k; else if (x.role==='Teacher'||x.role==='User') byKey['T|'+(x.school||'')+'|'+nn(x.name)] = k; });
    const roles = ['Admin','SchoolAdmin','Supervisor','Teacher','User'];
    (p.users||[]).forEach(x=>{
      let un = String(x.username||'').trim();
      if (un && !curU[un]){ const mk = x.role==='Supervisor' ? byKey['S|'+nn(x.name)] : byKey['T|'+(x.school||'')+'|'+nn(x.name)]; if (mk){ un = mk; x.password = ''; } }
      if (!un || !x.name){ skip.push(un||'(ไม่มี username)'); return; }
      if (un===u.username && x.role && x.role!=='Admin'){ skip.push(un+' (ห้ามลดสิทธิ์ตนเอง)'); return; }
      const ex = curU[un];
      const d = { username:un, name:String(x.name).trim(), position:x.position||'', rank:x.rank||'', group:x.group||'',
        school:x.school||'', schools:x.schools||[], role: roles.indexOf(x.role)>=0 ? x.role : 'Teacher',
        status: x.status==='suspended' ? 'suspended' : 'active', updated_at:now, updated_by:u.username };
      if (x.password) d.pass = hash_(String(x.password));
      else if (!ex) d.pass = hash_(un);
      if (!ex){ d.signature_url=''; uNew++; } else uUpd++;
      uOps.push({op:'merge', id:un, data:d});
    });
    if (uOps.length) writeColl_('Users', uOps);
    return JSON.stringify({ ok:true, sNew, sUpd, uNew, uUpd, skip, data: p.final ? loadFor_(u) : null });
  } finally { lock.releaseLock(); }
}

/* ---------------- สร้างบัญชีครูตามชื่อย่อโรงเรียน + รหัสผ่านสุ่ม ----------------
   payload {school, abbr}  →  username = ABBR_001, ABBR_002 ...  (รหัสผ่านส่งกลับครั้งเดียว ไม่เก็บแบบอ่านได้) */
function apiGenAccounts(token, json){
  const u = me_(token); need_(u,['Admin']);
  const p = JSON.parse(json), abbr = String(p.abbr||'').trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9]{1,9}$/.test(abbr)) throw new Error('ชื่อย่อต้องเป็นภาษาอังกฤษ/ตัวเลข 2–10 ตัว ขึ้นต้นด้วยตัวอักษร');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const users = readColl_('Users');
    const targets = Object.keys(users).map(k=>users[k])
      .filter(x=>x.school===p.school && (x.role==='Teacher'||x.role==='User'))
      .sort((a,b)=>String(a.username).localeCompare(String(b.username)));
    if (!targets.length) throw new Error('ไม่พบครูในโรงเรียนนี้');
    const tset = {}; targets.forEach(x=>tset[x.username]=1);
    const taken = {}; Object.keys(users).forEach(k=>{ if (!tset[k]) taken[k.toLowerCase()]=1; });
    const A = 'abcdefghjkmnpqrstuvwxyz23456789', rnd = n=>{ let r=''; for (let i=0;i<n;i++) r += A.charAt(Math.floor(Math.random()*A.length)); return r; };
    const ops = [], map = {}, out = []; let n = 1;
    const now = new Date().toISOString();
    targets.forEach(x=>{
      let nu; do { nu = abbr+'_'+('00'+n).slice(-3); n++; } while (taken[nu.toLowerCase()]);
      taken[nu.toLowerCase()] = 1;
      const pw = rnd(8);
      const d = Object.assign({}, x, { username:nu, pass:hash_(pw), updated_at:now, updated_by:u.username });
      if (nu!==x.username){ ops.push({op:'delete', id:x.username}); map[x.username]=nu; }
      ops.push({op:'set', id:nu, data:d});
      out.push({ name:x.name, position:x.position||'', rank:x.rank||'', role:x.role, username:nu, password:pw, old:x.username });
    });
    writeColl_('Users', ops);
    if (Object.keys(map).length){
      ['Evaluations','LessonPlans','Bookings'].forEach(c=>{
        const all = readColl_(c), mops = [];
        Object.keys(all).forEach(k=>{ const t = all[k].teacher_id; if (map[t]) mops.push({op:'merge', id:k, data:{teacher_id:map[t]}}); });
        if (mops.length) writeColl_(c, mops);
      });
    }
    const sc = readColl_('Schools'), sid = Object.keys(sc).find(k=>sc[k].name===p.school);
    if (sid) writeColl_('Schools', [{op:'merge', id:sid, data:{abbr:abbr}}]);
    return JSON.stringify({ ok:true, accounts:out, renamed:Object.keys(map).length, data:loadFor_(u) });
  } finally { lock.releaseLock(); }
}

/* ---------------- ล้างข้อมูล (Admin) ----------------
   mode 'evals' = ผลประเมิน + แผน + การจอง + วันงดนิเทศ
   mode 'all'   = evals + ผู้ใช้ทั้งหมด (ยกเว้นแอดมินระบบ) + โรงเรียน */
function apiClearData(token, mode){
  const u = me_(token); need_(u,['Admin']);
  if (mode!=='evals' && mode!=='all') throw new Error('โหมดไม่ถูกต้อง');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const wipe = name => { const sh = sh_(name), n = sh.getLastRow(); if (n>1) sh.getRange(2,1,n-1,3).clearContent(); };
    ['Evaluations','LessonPlans','Bookings'].forEach(wipe);
    writeColl_('Settings',[{op:'merge', id:'main', data:{locked_dates:[]}}]);
    if (mode==='all'){
      wipe('Schools'); wipe('Surveys');
      const us = readColl_('Users'), ops = [];
      Object.keys(us).forEach(k=>{ if (us[k].role!=='Admin') ops.push({op:'delete', id:k}); });
      if (ops.length) writeColl_('Users', ops);
    }
    return JSON.stringify({ ok:true, data: loadFor_(u) });
  } finally { lock.releaseLock(); }
}

/* ---------------- ลิงก์สั้นสำหรับแต่ละส่วนงาน ---------------- */
function apiShortLinks(token, mode, json){
  need_(me_(token),['Admin']);
  const props = PropertiesService.getScriptProperties(), base = ScriptApp.getService().getUrl();
  let cache = {}; try{ cache = JSON.parse(props.getProperty('SHORT_LINKS')||'{}'); }catch(e){}
  if (mode==='set'){
    const p = JSON.parse(json||'{}'), out = Object.assign({}, cache.base===base?cache:{}, {base:base});
    ['teacher','supervisor','schooladmin'].forEach(function(k){ if (p[k]!=null){ const v = String(p[k]).trim();
      if (v && !/^https:[/][/]\S+$/.test(v)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https:'); out[k] = v || (base+'?p='+k); } });
    props.setProperty('SHORT_LINKS', JSON.stringify(out)); return JSON.stringify(out);
  }
  const bad = function(v){ return !v || /tinyurl|is[.]gd|v[.]gd/i.test(v); };
  if (mode!=='make' && cache.base===base && !bad(cache.teacher) && !bad(cache.supervisor) && !bad(cache.schooladmin)) return JSON.stringify(cache);
  const out = { base: base };
  ['teacher','supervisor','schooladmin'].forEach(function(k){
    const long = base + '?p=' + k; let s = '';
    const tries = ['https://da.gd/s?url=', 'https://is.gd/create.php?format=simple&url='];
    for (let i=0;i<tries.length && !s;i++){
      try{ const r = UrlFetchApp.fetch(tries[i]+encodeURIComponent(long), {muteHttpExceptions:true, headers:{'User-Agent':'curl/8'}});
        const t = String(r.getContentText()).trim(); if (r.getResponseCode()===200 && /^https:[/][/]\S+$/.test(t)) s = t; }catch(e){}
    }
    out[k] = s || long;
  });
  props.setProperty('SHORT_LINKS', JSON.stringify(out));
  return JSON.stringify(out);
}

/* ---------------- ข้อมูลจำลองสำหรับทดลองระบบ (demo:true) ---------------- */
function apiDemo(token, mode, json){
  const u = me_(token); need_(u,['Admin']);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const COL = ['Users','Schools','Evaluations','LessonPlans','Bookings'];
    if (mode==='clear'){
      let n = 0;
      COL.forEach(function(c){ const all = readColl_(c), ops = [];
        Object.keys(all).forEach(function(k){ if (all[k] && all[k].demo===true && k!==u.username) ops.push({op:'delete', id:k}); });
        if (ops.length){ writeColl_(c, ops); n += ops.length; } });
      return JSON.stringify({ ok:true, n:n, data: loadFor_(u) });
    }
    if (mode!=='seed') throw new Error('โหมดไม่ถูกต้อง');
    const p = JSON.parse(json), exist = readColl_('Users');
    COL.forEach(function(c){ const list = p[c]||[]; if (!list.length) return;
      writeColl_(c, list.map(function(x){ const d = Object.assign({}, x.data, {demo:true});
        if (c==='Users'){ if (exist[x.id] && exist[x.id].demo!==true) return null; d.pass = hash_(d.password||'demo1234'); delete d.password; d.status = 'active'; }
        return {op:'set', id:x.id, data:d}; }).filter(Boolean)); });
    return JSON.stringify({ ok:true, data: loadFor_(u) });
  } finally { lock.releaseLock(); }
}

/* ================= ติดตั้งแบบ Template (เมนูในสเปรดชีต) ================= */
const ORIGINAL_SHEET_ID = '1rTClhpRyDxWid0EQB7kJf0NHTbmMdBjA3dGpvOwZdEI';
function isOriginal_(){ try{ return ss_().getId()===ORIGINAL_SHEET_ID; }catch(e){ return false; } }
function onOpen(){
  SpreadsheetApp.getUi().createMenu('🧭 ระบบนิเทศ')
    .addItem('① ติดตั้งระบบ (ครั้งแรก)', 'setupWizard')
    .addItem('② วิธีเผยแพร่เป็นเว็บแอป', 'showDeployHelp')
    .addItem('③ เปิด / คัดลอกลิงก์เว็บแอป', 'openWebApp')
    .addSeparator()
    .addItem('🔄 ตรวจเวอร์ชัน / วิธีอัปเดตระบบ', 'menuSelfUpdate')
    .addSeparator()
    .addItem('รีเซ็ตรหัสผ่านแอดมิน', 'resetAdminPassword')
    .addToUi();
}
function setupWizard(){
  const ui = SpreadsheetApp.getUi();
  if (isOriginal_()){ ui.alert('ไฟล์นี้คือระบบต้นฉบับที่ใช้งานจริง','เมนูติดตั้งใช้กับไฟล์ที่ทำสำเนาไปแล้วเท่านั้น จึงไม่ล้างข้อมูลในไฟล์นี้', ui.ButtonSet.OK); return; }
  const has = Object.keys(readColl_('Evaluations')).length + Object.keys(readColl_('Users')).length;
  if (has && ui.alert('ติดตั้งระบบใหม่?','ไฟล์นี้มีข้อมูลอยู่แล้ว การติดตั้งจะล้างข้อมูลทั้งหมดและเริ่มใหม่ ดำเนินการต่อหรือไม่', ui.ButtonSet.YES_NO)!==ui.Button.YES) return;
  const ask = function(t, m, def){ const r = ui.prompt(t, m+(def?'\n(เว้นว่าง = '+def+')':''), ui.ButtonSet.OK_CANCEL); if (r.getSelectedButton()!==ui.Button.OK) throw new Error('CANCEL'); return String(r.getResponseText()||'').trim() || def || ''; };
  try{
    const org  = ask('ขั้นที่ 1/4 · ชื่อหน่วยงาน', 'ชื่อที่แสดงบนหัวระบบและรายงาน', 'กลุ่มงานนิเทศการศึกษา');
    const aff  = ask('ขั้นที่ 2/4 · สังกัด', 'เช่น กองการศึกษา ศาสนา และวัฒนธรรม องค์การบริหารส่วนจังหวัด…', '');
    const user = ask('ขั้นที่ 3/4 · ชื่อผู้ใช้แอดมิน', 'ภาษาอังกฤษ ไม่เว้นวรรค', 'admin').toLowerCase().replace(/[^a-z0-9_.]/g,'') || 'admin';
    let pass = ''; while (pass.length < 6){ pass = ask('ขั้นที่ 4/4 · รหัสผ่านแอดมิน', 'อย่างน้อย 6 ตัวอักษร (จดไว้ให้ดี)', ''); if (pass.length<6) ui.alert('รหัสผ่านสั้นเกินไป'); }
    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try{
      COLLS.concat(['Surveys']).forEach(function(n){ const s = sh_(n), r = s.getLastRow(); if (r>1) s.getRange(2,1,r-1,3).clearContent(); });   /* รวมผลประเมินความพึงพอใจของต้นฉบับ ไม่ให้ติดไปกับสำเนา */
      const st = defaultSettings_(); st.school_info.name = org; st.school_info.affiliation = aff; st.school_info.logo = ''; st.school_info.schools = [];
      writeColl_('Settings', [{op:'set', id:'main', data: st}]);
      writeColl_('Users', [{op:'set', id:user, data:{ username:user, pass:hash_(pass), name:'ผู้ดูแลระบบ', role:'Admin', status:'active', school:'', group:'', signature_url:'' }}]);
      const p = PropertiesService.getScriptProperties(); p.deleteProperty('SHORT_LINKS'); p.setProperty('SETUP_DONE','1'); p.setProperty('SEED_V1','1'); p.setProperty('SEED_V2','1');
    } finally { lock.releaseLock(); }
    ui.alert('ติดตั้งเรียบร้อย ✅', 'แอดมิน: '+user+'\n\nขั้นต่อไป: เมนู 🧭 ระบบนิเทศ › ② วิธีเผยแพร่เป็นเว็บแอป', ui.ButtonSet.OK);
  }catch(e){ if (String(e.message)!=='CANCEL') throw e; }
}
function resetAdminPassword(){
  const ui = SpreadsheetApp.getUi(), us = readColl_('Users'), admins = Object.keys(us).filter(function(k){ return us[k].role==='Admin'; });
  if (!admins.length){ ui.alert('ไม่พบบัญชีแอดมิน — ใช้เมนู ① ติดตั้งระบบ'); return; }
  const r = ui.prompt('รีเซ็ตรหัสผ่าน', 'ชื่อผู้ใช้แอดมิน ('+admins.join(', ')+')', ui.ButtonSet.OK_CANCEL); if (r.getSelectedButton()!==ui.Button.OK) return;
  const id = String(r.getResponseText()).trim(); if (admins.indexOf(id)<0){ ui.alert('ไม่พบผู้ใช้นี้'); return; }
  const p = ui.prompt('รหัสผ่านใหม่', 'อย่างน้อย 6 ตัวอักษร', ui.ButtonSet.OK_CANCEL); if (p.getSelectedButton()!==ui.Button.OK) return;
  const pw = String(p.getResponseText()); if (pw.length<6){ ui.alert('รหัสผ่านสั้นเกินไป'); return; }
  writeColl_('Users', [{op:'merge', id:id, data:{pass:hash_(pw), status:'active'}}]); ui.alert('เปลี่ยนรหัสผ่านแล้ว');
}
function showDeployHelp(){
  const h = '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:14px;line-height:1.7;color:#1e293b">'
   + '<p style="margin-top:0">ทำครั้งเดียว ประมาณ 1 นาที</p><ol style="padding-left:20px">'
   + '<li>เมนู <b>ส่วนขยาย › Apps Script</b></li>'
   + '<li>มุมขวาบน กด <b>การทำให้ใช้งานได้ › การทำให้ใช้งานได้รายการใหม่</b></li>'
   + '<li>กดรูปเฟือง ⚙ เลือก <b>เว็บแอป</b></li>'
   + '<li>เรียกใช้ในฐานะ: <b>ฉัน</b> · ผู้ที่มีสิทธิ์เข้าถึง: <b>ทุกคน</b></li>'
   + '<li>กด <b>ทำให้ใช้งานได้</b> แล้วอนุญาตสิทธิ์ (ถ้ามีคำเตือน กด ขั้นสูง › ไปที่โครงการ)</li>'
   + '<li>กลับมาที่ชีตนี้ เมนู <b>🧭 ระบบนิเทศ › ③ เปิด / คัดลอกลิงก์เว็บแอป</b></li></ol>'
   + '<p style="background:#FFF7ED;padding:8px 10px;border-radius:8px">เมื่อแก้โค้ดภายหลัง ให้ <b>จัดการการทำให้ใช้งานได้ › ✏️ › เวอร์ชันใหม่</b> เพื่อคงลิงก์เดิม</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h).setWidth(460).setHeight(400), 'เผยแพร่เป็นเว็บแอป');
}
function openWebApp(){
  let url = ''; try{ url = ScriptApp.getService().getUrl() || ''; }catch(e){}
  const ui = SpreadsheetApp.getUi();
  if (!url){ ui.alert('ยังไม่ได้เผยแพร่เป็นเว็บแอป','ทำตามเมนู ② ก่อน', ui.ButtonSet.OK); return; }
  const h = '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:14px"><p>ลิงก์เว็บแอปของท่าน</p>'
   + '<input id="u" value="'+url+'" style="width:100%;padding:8px;font-size:13px" readonly onclick="this.select()">'
   + '<p style="margin-top:14px"><a href="'+url+'" target="_blank" style="background:#F97316;color:#fff;padding:9px 16px;border-radius:8px;text-decoration:none">เปิดระบบ</a> '
   + '<button onclick="var i=document.getElementById(\'u\');i.select();document.execCommand(\'copy\');this.textContent=\'คัดลอกแล้ว ✓\'" style="padding:8px 14px;border-radius:8px;border:1px solid #ccc;background:#fff">คัดลอกลิงก์</button></p>'
   + '<p style="color:#64748b;font-size:12px">เข้าระบบด้วยบัญชีแอดมินที่ตั้งไว้ แล้วเพิ่มโรงเรียน/ครูที่เมนู ข้อมูลโรงเรียน & นำเข้า</p></div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(h).setWidth(460).setHeight(220), 'เว็บแอประบบนิเทศ');
}

/* ---------------- ตรวจเวอร์ชัน / อัปเดตระบบ ----------------
   ดูเวอร์ชันล่าสุดจากศูนย์อัปเดต (GitHub Pages) · การอัปเดตทำโดยเจ้าของระบบ: วางโค้ด 2 ไฟล์ แล้ว deploy เวอร์ชันใหม่ (ระบบจึงทำงานในนามเจ้าของเสมอ) */
const UPD_BASE = ['https:', '', 'nonthaphathr.github.io', 'supervision-update', ''].join('/');
function updFetch_(url){ const r = UrlFetchApp.fetch(url + (url.indexOf('?')<0?'?':'&') + 't=' + Date.now(), {muteHttpExceptions:true});
  if (r.getResponseCode()!==200) throw new Error('ดึงข้อมูลจากศูนย์อัปเดตไม่สำเร็จ ('+r.getResponseCode()+')'); return r.getContentText(); }
function updCurrentVersion_(){ try{ const m = HtmlService.createHtmlOutputFromFile('index').getContent().match(/APP_VERSION = '([0-9.]+)'/); return m?m[1]:''; }catch(e){ return ''; } }
function updLatest_(){ return JSON.parse(updFetch_(UPD_BASE+'version.json')); }
function updNewer_(a, b){ const x=String(a||'0').split('.').map(Number), y=String(b||'0').split('.').map(Number);
  for (let i=0;i<Math.max(x.length,y.length);i++){ const d=(x[i]||0)-(y[i]||0); if (d) return d>0; } return false; }
function updEditorUrl_(){ return ['https:', '', 'script.google.com', 'home', 'projects', ScriptApp.getScriptId(), 'edit'].join('/'); }
function apiUpdateInfo(token){ need_(me_(token),['Admin']); const v = updLatest_(), cur = updCurrentVersion_();
  return JSON.stringify({ok:true, latest:v.version, date:v.date||'', notes:v.notes||[], current:cur, newer:updNewer_(v.version, cur), warn:deployWarn_(),
    center:UPD_BASE, editor:updEditorUrl_(), owner:sheetOwner_()}); }
function menuSelfUpdate(){
  const ui = SpreadsheetApp.getUi();
  let v; try{ v = updLatest_(); }catch(e){ ui.alert('ตรวจสอบเวอร์ชันไม่สำเร็จ', String(e.message||e), ui.ButtonSet.OK); return; }
  const cur = updCurrentVersion_(), nw = updNewer_(v.version, cur), esc = function(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;'); };
  const btn = function(href, t, pri){ return '<a href="'+href+'" target="_blank" style="display:inline-block;margin:4px 6px 4px 0;padding:8px 14px;border-radius:8px;text-decoration:none;'+(pri?'background:#F97316;color:#fff':'border:1px solid #F97316;color:#C2410C')+'">'+t+'</a>'; };
  const h = '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:14px;line-height:1.7;color:#1e293b">'
    + '<p style="margin-top:0">ใช้อยู่: <b>'+esc(cur||'-')+'</b> · ล่าสุด: <b style="color:#C2410C">'+esc(v.version)+'</b>'+(nw?' <span style="background:#FEF3C7;padding:1px 8px;border-radius:9px">มีเวอร์ชันใหม่</span>':' ✅ เป็นเวอร์ชันล่าสุดแล้ว')+'</p>'
    + (v.notes&&v.notes.length?'<ul style="margin:0 0 8px 18px;padding:0">'+v.notes.slice(0,5).map(function(n){ return '<li>'+esc(n)+'</li>'; }).join('')+'</ul>':'')
    + '<b>วิธีอัปเดต (ประมาณ 2 นาที)</b><ol style="padding-left:20px;margin:4px 0">'
    + '<li>เปิด <b>ศูนย์อัปเดต</b> › กดคัดลอก <b>Code.gs</b> › ไปที่หน้าแก้โค้ด › คลิกไฟล์ รหัส.gs › Ctrl+A › Ctrl+V</li>'
    + '<li>ทำแบบเดียวกันกับ <b>index.html</b> แล้วกดบันทึก 💾</li>'
    + '<li>การทำให้ใช้งานได้ › จัดการการทำให้ใช้งานได้ › ✏️ › เวอร์ชัน: <b>เวอร์ชันใหม่</b> › การทำให้ใช้งานได้</li></ol>'
    + btn(UPD_BASE, '1) เปิดศูนย์อัปเดต', true) + btn(updEditorUrl_(), '2) เปิดหน้าแก้โค้ด', false)
    + '<p style="font-size:12px;color:#64748b;margin-bottom:0">ต้องให้เจ้าของไฟล์ชีตนี้เป็นคนกด deploy เพื่อให้ไฟล์ทั้งหมดเก็บในไดรฟ์ของเจ้าของ</p></div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(h).setWidth(520).setHeight(470), '🔄 อัปเดตระบบ');
}

/* ---------------- Drive ---------------- */
function folder_(name, parent){ const it = parent.getFoldersByName(name); return it.hasNext()?it.next():parent.createFolder(name); }
function path_(p, dir){ String(p).split('/').map(function(x){return x.trim();}).filter(String).forEach(function(x){ dir = folder_(x, dir); }); return dir; }
function uniq_(dir, name){ var m = String(name).match(/^(.*?)(\.[^.]*)?$/), b = m[1], ext = m[2]||'', n = name, i = 2;
  const taken = function(x){ const it = dir.getFilesByName(x); while (it.hasNext()){ if (!it.next().isTrashed()) return true; } return false; };   /* ไม่นับไฟล์ในถังขยะ → สร้างซ้ำได้ชื่อเดิม */
  while (taken(n)) n = b+' ('+(i++)+')'+ext; return n; }
/* ---- โฟลเดอร์เก็บไฟล์: ตรึงด้วย ID ต่อชีต (กันระบบไปใช้โฟลเดอร์ชื่อซ้ำในไดรฟ์ของบัญชีอื่นที่กด deploy) ---- */
function runAs_(){ try{ return String(Session.getEffectiveUser().getEmail()||'').toLowerCase(); }catch(e){ return ''; } }
function sheetOwner_(){ try{ const o = DriveApp.getFileById(ss_().getId()).getOwner(); return o ? String(o.getEmail()||'').toLowerCase() : ''; }catch(e){ return ''; } }
function rootPin_(){ try{ const j = JSON.parse(PropertiesService.getScriptProperties().getProperty('ROOT_PIN')||'null'); return (j && j.sheet===ss_().getId()) ? j : null; }catch(e){ return null; } }
function deployMsg_(owner, me){ return 'ระบบกำลังทำงานในนามบัญชี '+(me||'อื่น')+' ไม่ใช่เจ้าของระบบ ('+(owner||'-')+') จึงไม่บันทึกไฟล์ เพื่อไม่ให้ไฟล์ไปอยู่ผิดไดรฟ์ — ให้เจ้าของระบบเปิด Apps Script › การทำให้ใช้งานได้ › จัดการการทำให้ใช้งานได้ › ✏️ › เวอร์ชันใหม่ › การทำให้ใช้งานได้'; }
function root_(){
  if (ROOT_FOLDER_ID) return DriveApp.getFolderById(ROOT_FOLDER_ID);
  const pin = rootPin_(), me = runAs_(), owner = sheetOwner_();
  /* Apps Script ทำงานในนามคนที่กด deploy ล่าสุดเสมอ ไฟล์ที่สร้างจะเป็นของบัญชีนั้นและกินพื้นที่บัญชีนั้น → ถ้าไม่ใช่เจ้าของชีต ไม่ยอมเขียนไฟล์ */
  if (owner && me && owner!==me) throw new Error(deployMsg_(owner, me));
  if (pin){ let f = null; try{ f = DriveApp.getFolderById(pin.id); }catch(e){}
    if (f && !f.isTrashed()) return f; }
  const f = folder_('ระบบนิเทศ-ไฟล์', DriveApp.getRootFolder());
  PropertiesService.getScriptProperties().setProperty('ROOT_PIN', JSON.stringify({sheet:ss_().getId(), id:f.getId(), owner:me, at:new Date().toISOString()}));
  return f;
}
function deployWarn_(){ const owner = sheetOwner_(), me = runAs_(); return (owner && me && owner!==me) ? {owner:owner, runAs:me, msg:deployMsg_(owner, me)} : null; }
function apiRootFolderUrl(token){ need_(me_(token),['Admin','Supervisor']); return JSON.stringify({ok:true, url: root_().getUrl()}); }

// payload {filename, mimeType, data(base64), section, sub}
function apiUpload(token, json){
  const u = me_(token);
  const p = JSON.parse(json);
  let dir = root_();
  if (p.year) dir = yearDir_(p.year);
  if (p.section) dir = path_(p.section, dir);
  if (p.sub) dir = path_(p.sub, dir);
  const file = dir.createFile(Utilities.newBlob(Utilities.base64Decode(p.data), p.mimeType, uniq_(dir, p.filename)));
  try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  try{ file.setDescription('nithet-upload:'+u.username); }catch(e){}
  const id = file.getId();
  try{ CacheService.getScriptCache().put('up_'+id, u.username, 21600); }catch(e){}
  const url = /^image\//.test(p.mimeType) ? 'https://lh3.googleusercontent.com/d/'+id+'=w1600'
                                           : 'https://drive.google.com/file/d/'+id+'/view';
  return JSON.stringify({ ok:true, id, url, by:u.username });
}
/* ---------- แทนที่ไฟล์: ไฟล์เก่าที่ไม่มีใครอ้างอิงแล้วย้ายลงถังขยะ ---------- */
const PLAN_FILE_KEYS_ = ['pdf_plan_url','post_note_url'];
function driveId_(url){ const m = String(url||'').match(/\/d\/([A-Za-z0-9_-]{20,})/) || String(url||'').match(/[?&]id=([A-Za-z0-9_-]{20,})/); return m ? m[1] : ''; }
function isUnder_(item, targetId, depth){
  const it = item.getParents();
  while (it.hasNext()){ const p = it.next(); if (p.getId()===targetId) return true; if (depth<8 && isUnder_(p, targetId, depth+1)) return true; }
  return false;
}
function referenced_(id){
  return ['LessonPlans','Evaluations','Users','Settings','Schools'].some(function(c){ return JSON.stringify(readColl_(c)).indexOf(id)>=0; });
}
function planFileIds_(old){
  const ids = PLAN_FILE_KEYS_.map(function(f){ return driveId_(old[f]); });
  const t = old.tfiles; if (t && typeof t==='object') Object.keys(t).forEach(function(k){ ids.push(driveId_(t[k])); });
  return ids.filter(function(x){ return x; });
}
function trashReplacedPlanFiles_(u, prev, ops){
  try{
    ops.forEach(function(o){ const old = prev[o.id]; if (!old) return;
      planFileIds_(old).forEach(function(id){ if (referenced_(id)) return;
        try{
          const file = DriveApp.getFileById(id), d = String(file.getDescription()||'');
          if (d.indexOf('nithet-upload:')===0 && d!=='nithet-upload:'+old.teacher_id && u.role!=='Admin') return;
          if (underPlans_(file)) file.setTrashed(true);
        }catch(e){}
      });
    });
  }catch(e){}
}
/* แอดมินลบช่องอัปโหลดฝั่งครู: เอาไฟล์ของช่องนั้นออกจากแผน แล้วย้ายไฟล์ลงถังขยะ */
function purgeRemovedTeacherBlocks_(u, stPrev, ops){
  try{
    const o = ops.filter(function(x){ return x.id==='main' && x.data && Array.isArray(x.data.teacher_files); }).pop(); if (!o) return;
    const keep = o.data.teacher_files.map(function(b){ return b && b.id; });
    const gone = (stPrev.teacher_files||[]).map(function(b){ return b && b.id; }).filter(function(id){ return /^tf/.test(String(id||'')) && keep.indexOf(id)<0; });
    if (!gone.length) return;
    const lp = readColl_('LessonPlans'), w = [], ids = [];
    Object.keys(lp).forEach(function(k){ const t = lp[k].tfiles; if (!t) return; let ch = false; const nt = Object.assign({}, t), nn = Object.assign({}, lp[k].tfnames||{});
      gone.forEach(function(g){ if (nt[g]!=null){ const d = driveId_(nt[g]); if (d) ids.push(d); delete nt[g]; delete nn[g]; ch = true; } });
      if (ch) w.push({op:'merge', id:k, data:{tfiles:nt, tfnames:nn}}); });
    if (w.length) writeColl_('LessonPlans', w);
    ids.forEach(function(id){ if (referenced_(id)) return; try{ const f = DriveApp.getFileById(id); if (underPlans_(f)) f.setTrashed(true); }catch(e){} });
  }catch(e){}
}
/* ไฟล์ที่เพิ่งอัปโหลดแล้วกด ✕ หรืออัปโหลดทับก่อนบันทึก: ลบได้เฉพาะไฟล์ของตนเองที่ยังไม่ถูกบันทึกใช้ */
function apiDiscardUpload(token, id){
  const u = me_(token); id = String(id||'');
  const owner = CacheService.getScriptCache().get('up_'+id);
  if (!id || owner!==u.username || referenced_(id)) return JSON.stringify({ ok:false });
  try{ DriveApp.getFileById(id).setTrashed(true); }catch(e){}
  return JSON.stringify({ ok:true });
}
function underPlans_(file){
  const rootId = root_().getId(); let cur = file, saw = false;
  for (let d=0; d<12; d++){ const it = cur.getParents(); if (!it.hasNext()) return false; cur = it.next();
    if (cur.getName()==='แผนการจัดการเรียนรู้') saw = true; if (cur.getId()===rootId) return saw; }
  return false;
}
/* ---------- โฟลเดอร์แยกตามปีการศึกษา ---------- */
function yearDir_(y){ const n = String(y||'').replace(/[^0-9]/g,''); return n ? folder_('ปีการศึกษา '+n, root_()) : root_(); }
function mergeInto_(src, dst){
  const fi = src.getFiles(); while (fi.hasNext()) fi.next().moveTo(dst);
  const fo = src.getFolders(); while (fo.hasNext()){ const s = fo.next(), ex = dst.getFoldersByName(s.getName()); if (ex.hasNext()){ mergeInto_(s, ex.next()); s.setTrashed(true); } else s.moveTo(dst); }
}
/* ย้ายโครงสร้างเดิม (ประเภท/ภาคเรียนที่ X ปีการศึกษา Y) → ปีการศึกษา Y/ประเภท/ภาคเรียนที่ X — ทำครั้งเดียว */
function migrateFoldersV2_(){
  const pr = PropertiesService.getScriptProperties(); if (pr.getProperty('FOLDERS_V2')) return;
  try{
    const root = root_();
    ['รูปนิเทศ','แผนการจัดการเรียนรู้','รายงาน PDF'].forEach(function(sec){
      const it = root.getFoldersByName(sec); if (!it.hasNext()) return; const sf = it.next();
      const subs = []; const fo = sf.getFolders(); while (fo.hasNext()) subs.push(fo.next());
      subs.forEach(function(f){ const m = f.getName().match(/^ภาคเรียนที่\s*(\S+)\s*ปีการศึกษา\s*(\d+)/); if (!m) return;
        const dst = folder_(sec, yearDir_(m[2])), name = 'ภาคเรียนที่ '+m[1], ex = dst.getFoldersByName(name);
        if (ex.hasNext()){ mergeInto_(f, ex.next()); f.setTrashed(true); } else { f.moveTo(dst); f.setName(name); } });
      if (!sf.getFiles().hasNext() && !sf.getFolders().hasNext()) sf.setTrashed(true);
    });
    pr.setProperty('FOLDERS_V2','1');
  }catch(e){}
}
/* ---------- รายงานพื้นที่จัดเก็บ ---------- */
function apiStorageUsage(token){
  need_(me_(token),['Admin']);
  const st = getSettings_(), y = String(st.academic_year||'').replace(/[^0-9]/g,''), t = String(st.term||'');
  migrateFoldersV2_();
  const root = root_();
  const walk = function(folder, acc){ const fi = folder.getFiles(); while (fi.hasNext()){ acc.bytes += fi.next().getSize(); acc.files++; }
    const fo = folder.getFolders(); while (fo.hasNext()) walk(fo.next(), acc); return acc; };
  const years = [], it = root.getFolders();
  while (it.hasNext()){ const f = it.next(), m = f.getName().match(/^ปีการศึกษา\s*(\d+)/); if (m){ const a = walk(f, {bytes:0, files:0}); years.push({year:m[1], bytes:a.bytes, files:a.files, url:f.getUrl()}); } }
  years.sort(function(a,b){ return b.year.localeCompare(a.year); });
  const term = {bytes:0, files:0}, yf = root.getFoldersByName('ปีการศึกษา '+y);
  if (yf.hasNext()){ const secs = yf.next().getFolders(); while (secs.hasNext()){ const tf = secs.next().getFoldersByName('ภาคเรียนที่ '+t); while (tf.hasNext()) walk(tf.next(), term); } }
  let used = 0, limit = 0; try{ used = DriveApp.getStorageUsed(); limit = DriveApp.getStorageLimit(); }catch(e){}
  return JSON.stringify({ ok:true, year:y, term:t, termUsage:term, system:walk(root, {bytes:0, files:0}), years:years, rootUrl:root.getUrl(), driveUsed:used, driveLimit:limit, runAs:runAs_() });
}
/* คู่มือแอดมินโรงเรียน (PDF) สำหรับแนบท้ายเอกสารรหัส */
const SA_MANUAL_ID = '1ytFHrNUABPzZj1nMGNIk62sE7OOgt96v';
function apiSaManual(token){
  need_(me_(token),['Admin']);
  let blob = null;
  try{ blob = DriveApp.getFileById(SA_MANUAL_ID).getBlob(); }catch(e){}
  if (!blob){ try{ const r = UrlFetchApp.fetch('https://drive.google.com/uc?export=download&id='+SA_MANUAL_ID, {followRedirects:true, muteHttpExceptions:true});
    if (r.getResponseCode()===200 && String(r.getHeaders()['Content-Type']||'').indexOf('pdf')>=0) blob = r.getBlob(); }catch(e){} }
  if (!blob) return JSON.stringify({ ok:false });
  return JSON.stringify({ ok:true, data: Utilities.base64Encode(blob.getBytes()) });
}
function apiDeleteFiles(token, idsJson){
  need_(me_(token),['Admin','Supervisor']);
  const done=[]; JSON.parse(idsJson).forEach(id=>{ try{ DriveApp.getFileById(id).setTrashed(true); done.push(id);}catch(e){} });
  return JSON.stringify({ok:true, deleted:done});
}

/* โลโก้: ส่งไฟล์ต้นฉบับเป็น data URL (คงพื้นโปร่งใสของ PNG ไม่ติดพื้นดำจากภาพย่อ) */
function logoData_(src){
  try{ const l = String(src || (getSettings_().school_info||{}).logo || '');
    const id = (l.match(/\/d\/([\w-]{10,})/) || l.match(/[?&]id=([\w-]{10,})/) || [])[1];
    if (!id) return '';
    const b = DriveApp.getFileById(id).getBlob(), by = b.getBytes();
    if (by.length > 700000) return '';
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(by);
  }catch(e){ return ''; }
}
function apiLogoData(token){ me_(token); return JSON.stringify({ok:true, data:logoData_('')}); }

/* ---------------- PDF report (HTML -> PDF in Drive) ---------------- */
// html สร้างจากหน้าเว็บ (มีตรา/ลายเซ็นเป็น URL) → แปลงรูปเป็น base64 ฝั่งเซิร์ฟเวอร์เพื่อให้ติดใน PDF
function apiMakePdf(token, json){
  const u = me_(token); need_(u,['Admin','Supervisor']);
  const p = JSON.parse(json);
  let html = String(p.html||'');
  html = html.replace(/src="(https:\/\/[^"]+)"/g, function(m, url){
    try{
      const id = (url.match(/\/d\/([\w-]{10,})/)||[])[1];
      const blob = id ? DriveApp.getFileById(id).getBlob() : UrlFetchApp.fetch(url).getBlob();
      return 'src="data:'+blob.getContentType()+';base64,'+Utilities.base64Encode(blob.getBytes())+'"';
    }catch(e){ return m; }
  });
  const dir = path_(p.sub||'อื่น ๆ', folder_('รายงาน PDF', p.year ? yearDir_(p.year) : root_()));
  // สร้างซ้ำ = แทนที่ไฟล์เดิมของการประเมินนี้ (ย้ายไฟล์เก่าไปถังขยะ แล้วใช้ชื่อเดิม)
  if (p.eval_id){ try{ const ev = (readColl_('Evaluations')||{})[p.eval_id]||{};
    const oid = (String(ev.pdf_drive_link||'').match(/\/d\/([\w-]{10,})/)||[])[1];
    if (oid) DriveApp.getFileById(oid).setTrashed(true); }catch(e){} }
  const name = uniq_(dir, (p.filename||'report')+'.pdf');
  const blob = p.pdf ? Utilities.newBlob(Utilities.base64Decode(p.pdf), 'application/pdf', name)
                     : Utilities.newBlob(html, 'text/html', 'r.html').getAs('application/pdf').setName(name);
  const f = dir.createFile(blob);
  try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  const url = 'https://drive.google.com/file/d/'+f.getId()+'/view';
  if (p.eval_id){ writeColl_('Evaluations',[{op:'merge', id:p.eval_id, data:{pdf_drive_link:url}}]); }
  return JSON.stringify({ok:true, url});
}

/* ---------------- ครั้งแรก: กดรันเพื่อขอสิทธิ์ ---------------- */
function authorizeOnce(){
  const f = DriveApp.getRootFolder().createFile('test.txt','x');
  try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  f.setTrashed(true);
  UrlFetchApp.fetch('https://www.google.com');
  PropertiesService.getScriptProperties().deleteProperty('SETUP_DONE');
  ensureSetup_();
  Logger.log('พร้อมใช้งาน — ผู้ใช้เริ่มต้น admin / admin1234 (เปลี่ยนรหัสทันที)');
}

/* ---------------- ประเมินความพึงพอใจ (ไม่ระบุตัวตน) ----------------
   ชีต Surveys: ไม่เก็บ username และไม่เก็บเวลาบันทึก · แทรกแถวตำแหน่งสุ่มเพื่อไม่ให้ลำดับแถวเผยลำดับเวลา
   ไม่อยู่ใน COLLS จึงเขียน/ลบผ่าน apiCommit ไม่ได้ — เขียนได้เฉพาะสองฟังก์ชันด้านล่าง */
function svRound_(){ const s = getSettings_(); return String(s.academic_year)+'/'+String(s.term); }
function svClean_(sv, p){
  const items = (sv.items||[]).filter(x=>x.type!=='section'), byId = {}; items.forEach(x=>{ byId[x.id]=x; });
  const ans = {}, a = (p&&p.answers)||{};
  Object.keys(a).forEach(k=>{ const q = byId[k]; if (!q) return; let v = a[k];
    if (q.type==='rating'){ v = +v; if (v>=1 && v<=5) ans[k] = Math.round(v); }
    else if (q.type==='single'){ v = String(v==null?'':v).slice(0,200); if (v && (q.options||[]).indexOf(v)>=0) ans[k] = v; }
    else if (q.type==='multi'){ v = (Array.isArray(v)?v:[]).map(x=>String(x).slice(0,200)).filter(x=>(q.options||[]).indexOf(x)>=0).slice(0,20); if (v.length) ans[k] = v; }
    else { v = String(v==null?'':v).trim().slice(0,1000); if (v) ans[k] = v; } });
  items.forEach(q=>{ if (q.required && ans[q.id]==null) throw new Error('กรุณาตอบข้อที่บังคับให้ครบ'); });
  if (!Object.keys(ans).length) throw new Error('ยังไม่ได้ตอบแบบประเมิน');
  return ans;
}
function svStore_(sv, p, via, group, school){
  if (!sv.open) throw new Error('ปิดรับการประเมินแล้ว');
  if (p && p.hp) throw new Error('ไม่สำเร็จ');
  const cache = CacheService.getScriptCache(), ck = 'svr'+Math.floor(Date.now()/60000), n = +(cache.get(ck)||0);
  if (n>=60) throw new Error('มีผู้ใช้งานจำนวนมาก กรุณาลองใหม่ในอีกสักครู่');
  cache.put(ck, String(n+1), 120);
  const rec = { round:svRound_(), via:via, group:group, school:school||'', fv:+sv.v||1, a:svClean_(sv, p) };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const s = sh_('Surveys'), last = s.getLastRow(), row = [[uid_('SV'), JSON.stringify(rec), '']];
    if (last>1){ const at = 2+Math.floor(Math.random()*(last-1)); s.insertRowBefore(at); s.getRange(at,1,1,3).setValues(row); }
    else s.getRange(2,1,1,3).setValues(row);
  } finally { lock.releaseLock(); }
}
/* ลิงก์เก็บตก: ไม่ต้องล็อกอิน ตรวจจากรหัสลิงก์ + สถานะเปิดรับ */
function apiSurveyPublic(json){
  const p = JSON.parse(json), sv = getSettings_().survey||{};
  if (!sv.key || String(p.key)!==String(sv.key)) throw new Error('ลิงก์ไม่ถูกต้องหรือถูกยกเลิกแล้ว');
  if (['ครู','ผู้บริหาร'].indexOf(p.group)<0) throw new Error('กรุณาเลือกกลุ่มผู้ตอบ');
  const names = Object.values(readColl_('Schools')).map(x=>x.name);
  svStore_(sv, p, 'link', p.group, (sv.ask_school!==false && names.indexOf(p.school)>=0) ? p.school : '');
  return JSON.stringify({ ok:true });
}
/* ผ่านระบบ: จำกัด 1 ครั้ง/รอบ/บัญชี โดยจดธงไว้ที่ผู้ใช้ (แยกจากเนื้อหาคำตอบ) */
function apiSurveySubmit(token, json){
  const u = me_(token); need_(u,['Teacher','User']);
  const p = JSON.parse(json), sv = getSettings_().survey||{}, rd = svRound_(), cur = readColl_('Users')[u.username]||{};
  if ((cur.surveyed||{})[rd]) throw new Error('ท่านประเมินในรอบนี้แล้ว');
  svStore_(sv, p, 'system', u.role==='User' ? 'ผู้บริหาร' : 'ครู', sv.ask_school===false ? '' : (u.school||''));
  const done = Object.assign({}, cur.surveyed||{}); done[rd] = true;
  writeColl_('Users', [{op:'merge', id:u.username, data:{surveyed:done}}]);
  return JSON.stringify({ ok:true });
}
