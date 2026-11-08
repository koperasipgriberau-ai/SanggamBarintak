import {initializeApp} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {getAuth,signInWithEmailAndPassword,signOut,onAuthStateChanged} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {getFirestore,collection,getDocs,doc,getDoc,writeBatch} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import {firebaseConfig} from "./firebase-config.js";
const app=initializeApp(firebaseConfig),auth=getAuth(app),db=getFirestore(app);
const $=s=>document.querySelector(s),rp=n=>"Rp"+(n||0).toLocaleString("id-ID");
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const VIEWS=["ringkasan","anggota","impor"];let role=null;
function show(t){VIEWS.forEach(v=>$("#"+v).classList.toggle("hide",v!==t));document.querySelectorAll("#nav button[data-t]").forEach(b=>b.setAttribute("aria-selected",b.dataset.t===t));if(t==="ringkasan")loadRingkasan();if(t==="anggota")loadAnggota()}
document.querySelectorAll("#nav button[data-t]").forEach(b=>b.onclick=()=>show(b.dataset.t));
$("#out").onclick=()=>signOut(auth);
$("#go").onclick=async()=>{const m=$("#lmsg");m.classList.add("hide");
 try{await signInWithEmailAndPassword(auth,$("#em").value.trim(),$("#pw").value)}catch(e){m.textContent="Email atau kata sandi tidak cocok.";m.classList.remove("hide")}};
onAuthStateChanged(auth,async u=>{
 if(!u){role=null;$("#login").classList.remove("hide");$("#nav").classList.add("hide");VIEWS.forEach(v=>$("#"+v).classList.add("hide"));$("#who").textContent="Masuk untuk mengelola data koperasi.";return}
 const a=await getDoc(doc(db,"admins",u.uid));
 if(!a.exists()){await signOut(auth);const m=$("#lmsg");m.textContent="Akun ini belum terdaftar sebagai pengurus.";m.classList.remove("hide");return}
 role=a.data().role;$("#login").classList.add("hide");$("#nav").classList.remove("hide");
 $("#who").textContent=`${a.data().nama||u.email} · ${role}`;
 $("#imp").disabled=!["bendahara","sekretaris"].includes(role);show("ringkasan")});
async function loadRingkasan(){const [a,s]=await Promise.all([getDocs(collection(db,"anggota")),getDocs(collection(db,"simpanan"))]);
 const t={pokok:0,wajib:0,sukarela:0};s.forEach(d=>{const x=d.data();t[x.jenis]+=x.jumlah});
 $("#stats").innerHTML=[[a.size,"anggota"],[rp(t.pokok),"simpanan pokok"],[rp(t.wajib),"simpanan wajib"],[rp(t.sukarela),"simpanan sukarela"],[rp(t.pokok+t.wajib+t.sukarela),"total simpanan"]].map(([b,l])=>`<div class="stat"><b>${b}</b><span>${l}</span></div>`).join("")}
async function loadAnggota(){const s=await getDocs(collection(db,"anggota"));const r=s.docs.map(d=>d.data()).sort((a,b)=>a.no-b.no);
 $("#ab").innerHTML=r.map(a=>`<tr><td class="n">${a.no}</td><td>${esc(a.nama)}</td><td>${esc(a.instansi)}</td><td>${esc(a.tglDaftar)}</td><td>${esc(a.status)}</td></tr>`).join("")||`<tr><td colspan="5">Belum ada anggota. Gunakan menu Impor Excel.</td></tr>`}
$("#f").onchange=()=>{};
const ymd=d=>d instanceof Date?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`:"";
$("#imp").onclick=async()=>{const m=$("#imsg");m.classList.remove("hide","err");const file=$("#f").files[0];if(!file){m.textContent="Pilih file Excel dulu.";return}
 $("#imp").disabled=true;m.textContent="Membaca file…";
 try{const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true});const ws=wb.Sheets["SIMPANAN ANGGOTA"];if(!ws)throw new Error('Sheet "SIMPANAN ANGGOTA" tidak ditemukan');
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null}).filter(r=>typeof r[0]==="number"&&typeof r[2]==="string"&&r[2].trim());
  const ops=[],pub=[];const TAHUN=2026;let tot={pokok:0,wajib:0,sukarela:0};
  for(const r of rows){const no=r[0],id="A"+String(no).padStart(3,"0"),tgl=ymd(r[1]);
   const pokok=+r[4]||0,suk=+r[5]||0;let wajib=0;
   ops.push([doc(db,"anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),tglDaftar:tgl,status:"aktif"}]);
   if(pokok)ops.push([doc(db,"simpanan",id+"_pokok"),{anggotaId:id,jenis:"pokok",jumlah:pokok,tanggal:tgl}]);
   if(suk)ops.push([doc(db,"simpanan",id+"_sukarela"),{anggotaId:id,jenis:"sukarela",jumlah:suk,tanggal:tgl}]);
   for(let b=0;b<12;b++){const v=+r[6+b]||0;if(v){wajib+=v;ops.push([doc(db,"simpanan",`${id}_wajib_${TAHUN}_${String(b+1).padStart(2,"0")}`),{anggotaId:id,jenis:"wajib",jumlah:v,periode:`${TAHUN}-${String(b+1).padStart(2,"0")}`}])}}
   tot.pokok+=pokok;tot.wajib+=wajib;tot.sukarela+=suk;
   ops.push([doc(db,"publik_anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),pokok,wajib,sukarela:suk,total:pokok+wajib+suk}])}
  ops.push([doc(db,"publik_ringkasan","utama"),{jumlahAnggota:rows.length,...tot,diperbarui:new Date().toISOString()}]);
  ops.push([doc(db,"config","shu"),{cadangan:40,pendidikanSosial:5,pengurusPengawas:10,shuAnggota:20,transaksiAnggota:15,simpananSukarela:10,catatan:"Cadangan 40% masuk kas cadangan; lima pos lainnya dibagikan dari SHU."}]);
  for(let i=0;i<ops.length;i+=400){const b=writeBatch(db);ops.slice(i,i+400).forEach(([ref,d])=>b.set(ref,d,{merge:true}));await b.commit();m.textContent=`Menyimpan… ${Math.min(i+400,ops.length)}/${ops.length}`}
  m.textContent=`Selesai: ${rows.length} anggota dan ${ops.length} catatan tersimpan.`}
 catch(e){m.classList.add("err");m.textContent="Impor gagal: "+e.message}
 $("#imp").disabled=!["bendahara","sekretaris"].includes(role)};
