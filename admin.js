import {initializeApp,deleteApp} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {getAuth,signInWithEmailAndPassword,signOut,onAuthStateChanged,createUserWithEmailAndPassword,sendPasswordResetEmail} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {getFirestore,collection,getDocs,doc,getDoc,setDoc,addDoc,writeBatch,serverTimestamp} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import {firebaseConfig} from "./firebase-config.js";
const app=initializeApp(firebaseConfig),auth=getAuth(app),db=getFirestore(app);
const $=s=>document.querySelector(s),rp=n=>"Rp"+(n||0).toLocaleString("id-ID");
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const MOD=[["anggota","Anggota"],["simpanan","Simpanan"],["bku","Buku kas (BKU)"],["surat","Surat masuk/keluar"],["shu","SHU"],["log","Log aktivitas"]];
const LBL=Object.fromEntries(MOD);
const VIEWS=["ringkasan","anggota","impor","pengguna"];
let me=null,USERS=[],editing=null;
const can=(m,k)=>!!me&&(me.master||(me[k]||[]).includes(m));
const ALLOWED={ringkasan:()=>can("anggota","lihat")&&can("simpanan","lihat"),anggota:()=>can("anggota","lihat"),impor:()=>can("anggota","ubah")&&can("simpanan","ubah"),pengguna:()=>me.master};
const msg=(el,t,err)=>{el.textContent=t;el.classList.remove("hide");el.classList.toggle("err",!!err)};
async function catat(aksi,detail){try{await addDoc(collection(db,"log"),{uid:auth.currentUser.uid,nama:me.nama,aksi,detail,waktu:serverTimestamp()})}catch(e){console.warn("log gagal",e)}}

function show(t){VIEWS.forEach(v=>$("#"+v).classList.toggle("hide",v!==t));
 document.querySelectorAll("#nav button[data-t]").forEach(b=>b.setAttribute("aria-selected",b.dataset.t===t));
 ({ringkasan:loadRingkasan,anggota:loadAnggota,pengguna:()=>{resetForm();loadUsers()}})[t]?.()}
document.querySelectorAll("#nav button[data-t]").forEach(b=>b.onclick=()=>show(b.dataset.t));
$("#out").onclick=()=>signOut(auth);
$("#go").onclick=async()=>{$("#lmsg").classList.add("hide");
 try{await signInWithEmailAndPassword(auth,$("#em").value.trim(),$("#pw").value)}catch(e){msg($("#lmsg"),"Email atau kata sandi tidak cocok.",true)}};
onAuthStateChanged(auth,async u=>{
 if(!u){me=null;$("#login").classList.remove("hide");$("#nav").classList.add("hide");VIEWS.forEach(v=>$("#"+v).classList.add("hide"));$("#who").textContent="Masuk untuk mengelola data koperasi.";return}
 let d=null;try{d=await getDoc(doc(db,"admins",u.uid))}catch(e){}
 if(!d||!d.exists()||d.data().aktif===false){await signOut(auth);msg($("#lmsg"),"Akun ini belum aktif sebagai pengurus. Hubungi master admin.",true);return}
 const x=d.data();me={uid:u.uid,email:u.email,nama:x.nama||u.email,jabatan:x.jabatan||x.role||"",master:x.master===true,lihat:x.lihat||[],ubah:x.ubah||[]};
 $("#login").classList.add("hide");$("#nav").classList.remove("hide");
 $("#who").textContent=`${me.nama} · ${me.master?"master admin":me.jabatan}`;
 let first=null;document.querySelectorAll("#nav button[data-t]").forEach(b=>{const ok=ALLOWED[b.dataset.t]();b.classList.toggle("hide",!ok);if(ok&&!first)first=b.dataset.t});
 if(first)show(first);else{VIEWS.forEach(v=>$("#"+v).classList.add("hide"));$("#who").textContent+=" · belum ada hak akses, hubungi master admin"}});

async function loadRingkasan(){try{const [a,s]=await Promise.all([getDocs(collection(db,"anggota")),getDocs(collection(db,"simpanan"))]);
 const t={pokok:0,wajib:0,sukarela:0};s.forEach(d=>{const x=d.data();t[x.jenis]+=x.jumlah});
 $("#stats").innerHTML=[[a.size,"anggota"],[rp(t.pokok),"simpanan pokok"],[rp(t.wajib),"simpanan wajib"],[rp(t.sukarela),"simpanan sukarela"],[rp(t.pokok+t.wajib+t.sukarela),"total simpanan"]].map(([b,l])=>`<div class="stat"><b>${b}</b><span>${l}</span></div>`).join("")}catch(e){$("#stats").innerHTML=`<div class="msg err">Data tidak bisa dimuat: ${esc(e.message)}</div>`}}
async function loadAnggota(){try{const s=await getDocs(collection(db,"anggota"));const r=s.docs.map(d=>d.data()).sort((a,b)=>a.no-b.no);
 $("#ab").innerHTML=r.map(a=>`<tr><td class="n">${a.no}</td><td>${esc(a.nama)}</td><td>${esc(a.instansi)}</td><td>${esc(a.tglDaftar)}</td><td>${esc(a.status)}</td></tr>`).join("")||`<tr><td colspan="5">Belum ada anggota. Gunakan menu Impor Excel.</td></tr>`}catch(e){$("#ab").innerHTML=`<tr><td colspan="5" class="msg err">Data tidak bisa dimuat: ${esc(e.message)}</td></tr>`}}

// ---------- Impor Excel ----------
const ymd=d=>d instanceof Date?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`:"";
$("#imp").onclick=async()=>{const m=$("#imsg");const file=$("#f").files[0];if(!file){msg(m,"Pilih file Excel dulu.",true);return}
 $("#imp").disabled=true;msg(m,"Membaca file…");
 try{const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true});const ws=wb.Sheets["SIMPANAN ANGGOTA"];if(!ws)throw new Error('Sheet "SIMPANAN ANGGOTA" tidak ditemukan');
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null}).filter(r=>typeof r[0]==="number"&&typeof r[2]==="string"&&r[2].trim());
  const ops=[];const TAHUN=2026;let tot={pokok:0,wajib:0,sukarela:0};
  for(const r of rows){const no=r[0],id="A"+String(no).padStart(3,"0"),tgl=ymd(r[1]);const pokok=+r[4]||0,suk=+r[5]||0;let wajib=0;
   ops.push([doc(db,"anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),tglDaftar:tgl,status:"aktif"}]);
   if(pokok)ops.push([doc(db,"simpanan",id+"_pokok"),{anggotaId:id,jenis:"pokok",jumlah:pokok,tanggal:tgl}]);
   if(suk)ops.push([doc(db,"simpanan",id+"_sukarela"),{anggotaId:id,jenis:"sukarela",jumlah:suk,tanggal:tgl}]);
   for(let b=0;b<12;b++){const v=+r[6+b]||0;if(v){wajib+=v;ops.push([doc(db,"simpanan",`${id}_wajib_${TAHUN}_${String(b+1).padStart(2,"0")}`),{anggotaId:id,jenis:"wajib",jumlah:v,periode:`${TAHUN}-${String(b+1).padStart(2,"0")}`}])}}
   tot.pokok+=pokok;tot.wajib+=wajib;tot.sukarela+=suk;
   ops.push([doc(db,"publik_anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),pokok,wajib,sukarela:suk,total:pokok+wajib+suk}])}
  ops.push([doc(db,"publik_ringkasan","utama"),{jumlahAnggota:rows.length,...tot,diperbarui:new Date().toISOString()}]);
  if(me.master)ops.push([doc(db,"config","shu"),{cadangan:40,pendidikanSosial:5,pengurusPengawas:10,shuAnggota:20,transaksiAnggota:15,simpananSukarela:10,catatan:"Cadangan 40% masuk kas cadangan; lima pos lainnya dibagikan dari SHU."}]);
  for(let i=0;i<ops.length;i+=400){const b=writeBatch(db);ops.slice(i,i+400).forEach(([ref,d])=>b.set(ref,d,{merge:true}));await b.commit();msg(m,`Menyimpan… ${Math.min(i+400,ops.length)}/${ops.length}`)}
  msg(m,`Selesai: ${rows.length} anggota dan ${ops.length} catatan tersimpan.`);await catat("impor_excel",`${rows.length} anggota`)}
 catch(e){msg(m,"Impor gagal: "+e.message,true)}
 $("#imp").disabled=false};

// ---------- Kelola pengguna (master) ----------
function buildIzin(sel){$("#izin").innerHTML=`<tr><th>Modul</th><th>Lihat</th><th>Ubah</th></tr>`+MOD.map(([m,l])=>`<tr><td>${l}</td>`+["lihat","ubah"].map(k=>`<td><input type="checkbox" style="width:auto" data-m="${m}" data-k="${k}" aria-label="${l}: ${k}" ${(sel[k]||[]).includes(m)?"checked":""}></td>`).join("")+`</tr>`).join("")}
$("#izin").onchange=e=>{const c=e.target;if(!c.dataset.m)return;const row=c.closest("tr"),l=row.querySelector("[data-k=lihat]"),u=row.querySelector("[data-k=ubah]");
 if(c.dataset.k==="ubah"&&c.checked)l.checked=true;if(c.dataset.k==="lihat"&&!c.checked)u.checked=false};
const collectIzin=()=>{const r={lihat:[],ubah:[]};document.querySelectorAll("#izin input").forEach(c=>{if(c.checked)r[c.dataset.k].push(c.dataset.m)});return r};
function resetForm(){editing=null;$("#ftitle").textContent="Tambah akun";["u_nama","u_email","u_pw"].forEach(i=>$("#"+i).value="");$("#u_jab").value="lainnya";
 $("#u_email").disabled=false;$("#pwwrap").classList.remove("hide");$("#u_aktif").checked=true;$("#u_aktif").disabled=false;$("#u_master").checked=false;$("#u_master").disabled=false;$("#u_mail").classList.add("hide");$("#u_msg").classList.add("hide");buildIzin({});syncMaster()}
function syncMaster(){$("#izin").classList.toggle("hide",$("#u_master").checked)}
$("#u_master").onchange=syncMaster;
$("#u_reset").onclick=resetForm;
const ringkasIzin=u=>u.master?"Semua (master)":(`Lihat: ${(u.lihat||[]).map(m=>LBL[m]||m).join(", ")||"-"} · Ubah: ${(u.ubah||[]).map(m=>LBL[m]||m).join(", ")||"-"}`);
async function loadUsers(){try{const s=await getDocs(collection(db,"admins"));USERS=s.docs.map(d=>({uid:d.id,...d.data()}));
 $("#ub").innerHTML=USERS.map(u=>`<tr><td>${esc(u.nama)}</td><td>${esc(u.jabatan||u.role||"")}</td><td>${esc(u.email||"")}</td><td>${u.aktif===false?"nonaktif":"aktif"}</td><td>${esc(ringkasIzin(u))}</td><td><button class="btn ghost" data-uid="${esc(u.uid)}">Ubah</button></td></tr>`).join("")}
 catch(e){$("#ub").innerHTML=`<tr><td colspan="6" class="msg err">Gagal memuat: ${esc(e.message)}</td></tr>`}}
$("#ub").onclick=e=>{const b=e.target.closest("button[data-uid]");if(!b)return;const u=USERS.find(x=>x.uid===b.dataset.uid);if(!u)return;
 editing=u;$("#ftitle").textContent="Ubah akun";$("#u_nama").value=u.nama||"";$("#u_jab").value=u.jabatan||u.role||"lainnya";$("#u_email").value=u.email||"";$("#u_email").disabled=true;
 $("#pwwrap").classList.add("hide");$("#u_aktif").checked=u.aktif!==false;$("#u_aktif").disabled=u.uid===me.uid;$("#u_master").checked=!!u.master;$("#u_master").disabled=u.uid===me.uid;buildIzin(u);syncMaster();
 $("#u_mail").classList.toggle("hide",!u.email);$("#u_msg").classList.add("hide");scrollTo({top:0,behavior:"smooth"})};
$("#u_mail").onclick=async()=>{try{await sendPasswordResetEmail(auth,editing.email);msg($("#u_msg"),"Email atur ulang kata sandi sudah dikirim ke "+editing.email)}catch(e){msg($("#u_msg"),"Gagal mengirim: "+e.message,true)}};
$("#u_save").onclick=async()=>{const m=$("#u_msg"),nama=$("#u_nama").value.trim(),jab=$("#u_jab").value,email=$("#u_email").value.trim(),izin=collectIzin();
 if(!nama){msg(m,"Isi nama terlebih dahulu.",true);return}
 $("#u_save").disabled=true;
 try{
  if(editing){const data={nama,jabatan:jab};if(editing.uid!==me.uid){data.aktif=$("#u_aktif").checked;data.master=$("#u_master").checked;data.lihat=izin.lihat;data.ubah=izin.ubah}
   await setDoc(doc(db,"admins",editing.uid),data,{merge:true});await catat("ubah_akun",`${nama}`);msg(m,"Perubahan tersimpan.")}
  else{const pw=$("#u_pw").value;
   if(!email){msg(m,"Isi email terlebih dahulu.",true);return}
   if(pw.length<12||!/[a-z]/.test(pw)||!/[A-Z]/.test(pw)||!/\d/.test(pw)){msg(m,"Kata sandi minimal 12 karakter dan memuat huruf besar, huruf kecil, dan angka.",true);return}
   const sec=initializeApp(firebaseConfig,"sec"+Date.now());let uid;
   try{const c=await createUserWithEmailAndPassword(getAuth(sec),email,pw);uid=c.user.uid;await signOut(getAuth(sec))}finally{await deleteApp(sec)}
   await setDoc(doc(db,"admins",uid),{nama,jabatan:jab,email,aktif:$("#u_aktif").checked,master:$("#u_master").checked,lihat:izin.lihat,ubah:izin.ubah});
   await catat("buat_akun",`${nama} (${email})`);msg(m,`Akun ${email} dibuat. Sampaikan kata sandi awalnya secara pribadi dan minta diganti setelah masuk.`);}
  await loadUsers();if(!editing)resetForm2(m)}
 catch(e){msg(m,e.code==="auth/email-already-in-use"?"Email ini sudah terdaftar.":"Gagal menyimpan: "+e.message,true)}
 $("#u_save").disabled=false};
function resetForm2(m){const t=m.textContent;resetForm();msg(m,t)}
