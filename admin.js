import {initializeApp,deleteApp} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {getAuth,signInWithEmailAndPassword,signOut,onAuthStateChanged,createUserWithEmailAndPassword,sendPasswordResetEmail} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {getFirestore,collection,getDocs,doc,getDoc,setDoc,addDoc,deleteDoc,deleteField,writeBatch,serverTimestamp,query,where,orderBy,limit} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import {firebaseConfig} from "./firebase-config.js";
const app=initializeApp(firebaseConfig),auth=getAuth(app),db=getFirestore(app);
const $=s=>document.querySelector(s),rp=n=>"Rp"+Math.round(n||0).toLocaleString("id-ID");
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const ymd=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const today=()=>ymd(new Date());
const msg=(el,t,err)=>{el.textContent=t;el.classList.remove("hide");el.classList.toggle("err",!!err)};
const unduh=(rows,nama)=>{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Data");XLSX.writeFile(wb,nama)};
const clean=o=>JSON.parse(JSON.stringify(o??null));
const MOD=[["anggota","Anggota"],["simpanan","Simpanan"],["bku","Buku kas (BKU)"],["surat","Surat masuk/keluar"],["shu","SHU"],["log","Log & riwayat"]];
const LBL=Object.fromEntries(MOD);
const KODE=["SIM","BIA","CAFE","KONVEKSI"];
const BLN=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
const BLNF=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const WAJIB=50000;
let me=null,ANG=[],SIM=[],BKU=[],SURAT=[],SHUROWS=[],USERS=[],editing=null,TUNG=[];
const can=(m,k)=>!!me&&(me.master||(me[k]||[]).includes(m));
async function getAnggota(force){if(force||!ANG.length){const s=await getDocs(collection(db,"anggota"));ANG=s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>a.no-b.no)}return ANG}

// ---------- Log & riwayat ----------
const logData=({aksi,kol="",id="",sebelum=null,sesudah=null,alasan="",detail=""})=>({uid:auth.currentUser.uid,nama:me.nama,aksi,koleksi:kol,docId:id,sebelum:clean(sebelum),sesudah:clean(sesudah),alasan,detail,waktu:serverTimestamp()});
const logOp=(b,x)=>b.set(doc(collection(db,"log")),logData(x));
async function catat(aksi,detail){try{await addDoc(collection(db,"log"),logData({aksi,detail}))}catch(e){console.warn("log gagal",e)}}
const fmt=v=>v===undefined||v===null||v===""?"(kosong)":typeof v==="object"?JSON.stringify(v):v;
const ring=o=>Object.entries(o||{}).filter(([k,v])=>typeof v!=="object").map(([k,v])=>`${k}=${v}`).join(", ");
function diffText(x){const a=x.sebelum,b=x.sesudah;
 if(a&&b)return Object.keys(b).filter(k=>JSON.stringify(a[k])!==JSON.stringify(b[k])).map(k=>`${k}: ${fmt(a[k])} → ${fmt(b[k])}`).join("; ")||x.detail||"";
 if(a&&!b)return "Data sebelum dihapus: "+ring(a);if(b&&!a)return "Data baru: "+ring(b);return x.detail||""}
const wkt=x=>x.waktu?.toDate?x.waktu.toDate().toLocaleString("id-ID"):"";
async function riwayat(kol,id){$("#dlg2_t").textContent="Riwayat perubahan";$("#dlg2_b").textContent="Memuat…";$("#dlg2").showModal();
 try{const s=await getDocs(query(collection(db,"log"),where("docId","==",id)));
  const r=s.docs.map(d=>d.data()).filter(x=>x.koleksi===kol).sort((a,b)=>(b.waktu?.seconds||9e9)-(a.waktu?.seconds||9e9));
  $("#dlg2_b").innerHTML=r.length?`<table><thead><tr><th>Waktu</th><th>Akun</th><th>Aksi</th><th>Rincian</th><th>Alasan</th></tr></thead><tbody>${r.map(x=>`<tr><td>${esc(wkt(x))}</td><td>${esc(x.nama)}</td><td>${esc(x.aksi)}</td><td>${esc(diffText(x))}</td><td>${esc(x.alasan||"")}</td></tr>`).join("")}</tbody></table>`:"<p>Belum ada riwayat untuk data ini. Data hasil impor Excel baru tercatat riwayatnya sejak diubah pertama kali.</p>"}
 catch(e){$("#dlg2_b").textContent="Gagal memuat riwayat: "+e.message}}
$("#dlg2_x").onclick=()=>$("#dlg2").close();

// ---------- Dialog isian ----------
function tanya({judul,fields,nilai={},validasi,alasan=true,tombol="Simpan"}){return new Promise(res=>{
 const d=$("#dlg");$("#dlg_t").textContent=judul;$("#dlg_ok").textContent=tombol;$("#dlg_e").classList.add("hide");
 $("#dlg_f").innerHTML=fields.map(f=>{if(f.type==="note")return `<p class="msg">${esc(f.label)}</p>`;const v=nilai[f.k]??"";
  const inp=f.type==="select"?`<select id="df_${f.k}" ${f.dis?"disabled":""}>${f.opt.map(([val,l])=>`<option value="${esc(val)}" ${String(val)===String(v)?"selected":""}>${esc(l)}</option>`).join("")}</select>`:`<input id="df_${f.k}" type="${f.type||"text"}" value="${esc(v)}" ${f.dis?"disabled":""}>`;
  return `<label for="df_${f.k}">${esc(f.label)}</label>${inp}`}).join("")+(alasan?`<label for="df_alasan">Alasan (wajib, tercatat di riwayat bersama nama akun Anda)</label><input id="df_alasan">`:"");
 const err=t=>{msg($("#dlg_e"),t,true)};
 $("#dlg_ok").onclick=()=>{const o={};fields.forEach(f=>{if(f.type==="note")return;const el=$("#df_"+f.k);o[f.k]=f.type==="number"?(el.value===""?0:+el.value):el.value.trim()});
  if(alasan){o.alasan=$("#df_alasan").value.trim();if(!o.alasan)return err("Isi alasan terlebih dahulu.")}
  for(const f of fields)if(f.req&&!o[f.k])return err(`${f.label} wajib diisi.`);
  const e=validasi?.(o);if(e)return err(e);res(o);d.close()};
 $("#dlg_x").onclick=()=>{res(null);d.close()};d.onclose=()=>res(null);d.showModal()})}
const btn=(a,id,l,k="")=>`<button class="btn ghost sm" data-a="${a}" data-id="${esc(id)}" data-k="${k}">${l}</button>`;
const rwBtn=(k,id)=>can("log","lihat")?btn("rw",id,"Riwayat",k):"";
const aksi=(sel,map)=>{$(sel).onclick=async e=>{const b=e.target.closest("button[data-a]");if(!b)return;b.disabled=true;try{await map[b.dataset.a](b.dataset.id,b)}catch(err){console.error(err);alert("Gagal: "+err.message)}b.disabled=false}};
const rw=(id,b)=>riwayat(b.dataset.k,id);

// ---------- Navigasi ----------
const VIEWS=[
 ["ringkasan","Ringkasan",()=>can("anggota","lihat")&&can("simpanan","lihat"),()=>loadRingkasan()],
 ["anggota","Anggota",()=>can("anggota","lihat"),()=>loadAnggota()],
 ["simpanan","Simpanan",()=>can("simpanan","lihat")&&can("anggota","lihat"),()=>loadSimpanan()],
 ["tunggakan","Tunggakan",()=>can("simpanan","lihat")&&can("anggota","lihat"),()=>loadTunggakan()],
 ["bku","Buku kas",()=>can("bku","lihat"),()=>loadBku()],
 ["surat","Surat",()=>can("surat","lihat"),()=>loadSurat()],
 ["shu","SHU",()=>can("shu","lihat")&&can("anggota","lihat")&&can("simpanan","lihat"),()=>loadShu()],
 ["impor","Impor Excel",()=>can("anggota","ubah")&&can("simpanan","ubah"),()=>{}],
 ["pengguna","Pengguna",()=>me.master,()=>{resetForm();loadUsers()}],
 ["log","Log",()=>can("log","lihat"),()=>loadLog()]];
async function show(t){VIEWS.forEach(v=>$("#"+v[0]).classList.toggle("hide",v[0]!==t));
 document.querySelectorAll("#tabs button").forEach(b=>b.setAttribute("aria-selected",b.dataset.t===t));
 document.querySelectorAll(".loaderr").forEach(e=>e.remove());
 try{await VIEWS.find(v=>v[0]===t)[3]()}catch(e){console.error(e);$("#"+t).insertAdjacentHTML("afterbegin",`<div class="msg err loaderr">Gagal memuat: ${esc(e.message)}</div>`)}}
$("#tabs").onclick=e=>{const b=e.target.closest("button[data-t]");if(b)show(b.dataset.t)};
$("#out").onclick=()=>signOut(auth);
$("#go").onclick=async()=>{$("#lmsg").classList.add("hide");
 try{await signInWithEmailAndPassword(auth,$("#em").value.trim(),$("#pw").value)}catch(e){msg($("#lmsg"),"Email atau kata sandi tidak cocok.",true)}};
onAuthStateChanged(auth,async u=>{
 if(!u){me=null;ANG=[];$("#login").classList.remove("hide");$("#nav").classList.add("hide");VIEWS.forEach(v=>$("#"+v[0]).classList.add("hide"));$("#who").textContent="Masuk untuk mengelola data koperasi.";return}
 let d=null;try{d=await getDoc(doc(db,"admins",u.uid))}catch(e){}
 if(!d||!d.exists()||d.data().aktif===false){await signOut(auth);msg($("#lmsg"),"Akun ini belum aktif sebagai pengurus. Hubungi master admin.",true);return}
 const x=d.data();me={uid:u.uid,email:u.email,nama:x.nama||u.email,jabatan:x.jabatan||x.role||"",master:x.master===true,lihat:x.lihat||[],ubah:x.ubah||[]};
 $("#login").classList.add("hide");$("#nav").classList.remove("hide");
 $("#who").textContent=`${me.nama} · ${me.master?"master admin":me.jabatan}`;
 const ok=VIEWS.filter(v=>v[2]());
 $("#tabs").innerHTML=ok.map(v=>`<button role="tab" data-t="${v[0]}" aria-selected="false">${v[1]}</button>`).join("");
 if(ok.length)show(ok[0][0]);else{VIEWS.forEach(v=>$("#"+v[0]).classList.add("hide"));$("#who").textContent+=" · belum ada hak akses, hubungi master admin"}});

// ---------- Sinkron data publik ----------
async function syncRingkasan(){await getAnggota();const aktif=new Set(ANG.filter(a=>a.status!=="berhenti").map(a=>a.id));
 const all=await getDocs(collection(db,"simpanan")),g={pokok:0,wajib:0,sukarela:0};all.forEach(d=>{const x=d.data();if(aktif.has(x.anggotaId))g[x.jenis]+=x.jumlah});
 await setDoc(doc(db,"publik_ringkasan","utama"),{jumlahAnggota:aktif.size,...g,diperbarui:new Date().toISOString()},{merge:true})}
async function syncPublik(id){const a=(await getAnggota()).find(x=>x.id===id);if(!a)return;
 if(a.status==="berhenti"){await deleteDoc(doc(db,"publik_anggota",id))}
 else{const q=await getDocs(query(collection(db,"simpanan"),where("anggotaId","==",id))),t={pokok:0,wajib:0,sukarela:0};q.forEach(d=>{const x=d.data();t[x.jenis]+=x.jumlah});
  await setDoc(doc(db,"publik_anggota",id),{no:a.no,nama:a.nama,instansi:a.instansi||"",...t,total:t.pokok+t.wajib+t.sukarela},{merge:true})}
 await syncRingkasan()}

// ---------- Tunggakan (dipakai juga oleh Ringkasan) ----------
function hitungTunggakan(ang,sim,tahun,jt){const n=new Date();let due=tahun<n.getFullYear()?12:tahun>n.getFullYear()?0:n.getMonth()+(n.getDate()>=jt?1:0);
 const paid={};sim.forEach(x=>{if(x.jenis==="wajib"&&x.periode&&x.periode.startsWith(tahun+"-"))(paid[x.anggotaId]??=new Set()).add(+x.periode.slice(5,7))});
 return ang.filter(a=>a.status!=="berhenti").map(a=>{const ty=a.tglDaftar?+a.tglDaftar.slice(0,4):0,mulai=ty>tahun?13:ty===tahun?+a.tglDaftar.slice(5,7):1,cells=[],tung=[];
  for(let m=1;m<=12;m++){if(paid[a.id]?.has(m))cells.push("bayar");else if(m<mulai||m>due)cells.push("-");else{cells.push("tunggak");tung.push(m)}}
  return {a,cells,tung,nominal:tung.length*WAJIB}})}
const normHp=h=>{let d=String(h||"").replace(/\D/g,"");if(!d)return"";if(d.startsWith("0"))d="62"+d.slice(1);else if(d.startsWith("8"))d="62"+d;return d};
const pesan=r=>`Assalamu'alaikum Bapak/Ibu ${r.a.nama}. Pengingat dari Koperasi PGRI Sanggam Barintak: simpanan wajib bulan ${r.tung.map(m=>BLNF[m-1]).join(", ")} (${rp(r.nominal)}) belum tercatat. Mohon ditransfer ke rekening Bank Kaltimtara 0062920998 a.n. KOPERASI PGRI SB. Abaikan pesan ini bila sudah membayar. Terima kasih.`;
async function loadTunggakan(){const [,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);SIM=s.docs.map(d=>({id:d.id,...d.data()}));
 if(!$("#t_thn").value)$("#t_thn").value=new Date().getFullYear();if(!$("#t_jt").value)$("#t_jt").value=10;renderTung()}
function renderTung(){const th=+$("#t_thn").value,jt=+$("#t_jt").value||10;TUNG=hitungTunggakan(ANG,SIM,th,jt);const p=TUNG.filter(r=>r.tung.length).sort((a,b)=>b.tung.length-a.tung.length);
 const tb=p.reduce((s,r)=>s+r.tung.length,0),tn=p.reduce((s,r)=>s+r.nominal,0);
 $("#t_cards").innerHTML=[[TUNG.length,"anggota aktif"],[p.length,"anggota menunggak"],[tb,"total bulan menunggak"],[rp(tn),"total nominal tunggakan"]].map(([b,l])=>`<div class="stat"><b>${b}</b><span>${l}</span></div>`).join("");
 $("#tb1").innerHTML=p.map((r,i)=>{const w=normHp(r.a.hp);return `<tr><td class="n">${i+1}</td><td>${esc(r.a.nama)}</td><td>${esc(r.a.instansi)}</td><td>${esc(r.a.hp||"")}</td><td>${r.tung.map(m=>BLN[m-1]).join(", ")}</td><td class="n">${r.tung.length}</td><td class="n">${rp(r.nominal)}</td><td class="noprint">${w?`<a class="btn ghost sm" target="_blank" rel="noopener" href="https://wa.me/${w}?text=${encodeURIComponent(pesan(r))}">WhatsApp</a>`:'<span class="dim">no HP belum ada</span>'}</td></tr>`}).join("")||`<tr><td colspan="8">Tidak ada tunggakan pada tahun ${th}.</td></tr>`;
 $("#th2").innerHTML=`<tr><th>Nama</th>${BLN.map(b=>`<th>${b}</th>`).join("")}<th class="n">Bulan menunggak</th></tr>`;
 $("#tb2").innerHTML=TUNG.map(r=>`<tr><td>${esc(r.a.nama)}</td>${r.cells.map(c=>c==="bayar"?'<td title="Sudah bayar">✓</td>':c==="tunggak"?'<td class="bad" title="Menunggak">✗</td>':'<td class="dim">-</td>').join("")}<td class="n">${r.tung.length}</td></tr>`).join("");
 $("#t_xls").onclick=()=>unduh(TUNG.map(r=>({Nama:r.a.nama,Instansi:r.a.instansi,HP:r.a.hp||"",...Object.fromEntries(BLN.map((b,i)=>[b,r.cells[i]==="bayar"?"Lunas":r.cells[i]==="tunggak"?"Menunggak":"-"])),"Bulan menunggak":r.tung.length,"Nominal tunggakan":r.nominal})),`rekap-simpanan-wajib-${th}.xlsx`);
 $("#t_copy").onclick=async()=>{const t=`Pengingat dari Koperasi PGRI Sanggam Barintak: simpanan wajib anggota Rp50.000 per bulan. Bagi Bapak/Ibu yang belum membayar, mohon ditransfer ke rekening Bank Kaltimtara 0062920998 a.n. KOPERASI PGRI SB. Abaikan pesan ini bila sudah membayar. Terima kasih.`;
  try{await navigator.clipboard.writeText(t);msg($("#t_msg"),"Pesan untuk grup sudah disalin. Tempel di grup WhatsApp.")}catch(e){msg($("#t_msg"),t)}}}
$("#t_thn").onchange=renderTung;$("#t_jt").onchange=renderTung;

// ---------- Ringkasan ----------
async function loadRingkasan(){const [a,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);SIM=s.docs.map(d=>({id:d.id,...d.data()}));
 const t={pokok:0,wajib:0,sukarela:0};SIM.forEach(x=>{t[x.jenis]+=x.jumlah});
 $("#stats").innerHTML=[[a.filter(x=>x.status!=="berhenti").length,"anggota aktif"],[rp(t.pokok),"simpanan pokok"],[rp(t.wajib),"simpanan wajib"],[rp(t.sukarela),"simpanan sukarela"],[rp(t.pokok+t.wajib+t.sukarela),"total simpanan"]].map(([b,l])=>`<div class="stat"><b>${b}</b><span>${l}</span></div>`).join("");
 const p=hitungTunggakan(a,SIM,new Date().getFullYear(),10).filter(r=>r.tung.length);
 $("#r_warn").innerHTML=p.length?`<div class="warn"><span><b>${p.length} anggota menunggak simpanan wajib</b> dengan total ${rp(p.reduce((s,r)=>s+r.nominal,0))}.</span><button class="btn sm" id="r_go">Lihat rekap & kirim pengingat</button></div>`:`<div class="warn"><span>Tidak ada tunggakan simpanan wajib tahun ini.</span></div>`;
 const g=$("#r_go");if(g)g.onclick=()=>show("tunggakan")}

// ---------- Anggota ----------
async function loadAnggota(){const r=await getAnggota(true),e=can("anggota","ubah");$("#a_form").classList.toggle("hide",!e);
 $("#ab").innerHTML=r.map(a=>`<tr><td class="n">${a.no}</td><td>${esc(a.nama)}</td><td>${esc(a.instansi)}</td><td>${esc(a.npa||"")}</td><td>${esc(a.hp||"")}</td><td>${esc(a.tglDaftar||"")}</td><td>${esc(a.status||"aktif")}</td><td>${e?btn("ubah",a.id,"Ubah"):""} ${rwBtn("anggota",a.id)}</td></tr>`).join("")||`<tr><td colspan="8">Belum ada anggota. Tambahkan di atas atau gunakan Impor Excel.</td></tr>`}
aksi("#ab",{rw,ubah:async id=>{const a=ANG.find(x=>x.id===id);
 const v=await tanya({judul:`Ubah anggota No. ${a.no}`,fields:[{k:"nama",label:"Nama",req:1},{k:"instansi",label:"Instansi / sekolah"},{k:"npa",label:"NPA PGRI"},{k:"hp",label:"No HP"},{k:"status",label:"Status",type:"select",opt:[["aktif","Aktif"],["berhenti","Berhenti"]]}],nilai:{...a,status:a.status||"aktif"}});if(!v)return;
 const sesudah={nama:v.nama,instansi:v.instansi,npa:v.npa,hp:v.hp,status:v.status},sebelum={nama:a.nama||"",instansi:a.instansi||"",npa:a.npa||"",hp:a.hp||"",status:a.status||"aktif"};
 if(JSON.stringify(sebelum)===JSON.stringify(sesudah))return alert("Tidak ada yang diubah.");
 const b=writeBatch(db);b.set(doc(db,"anggota",id),sesudah,{merge:true});logOp(b,{aksi:"ubah",kol:"anggota",id,sebelum,sesudah,alasan:v.alasan,detail:a.nama});await b.commit();
 await getAnggota(true);await syncPublik(id);await loadAnggota()}});
$("#a_save").onclick=async()=>{const m=$("#a_msg"),nama=$("#a_nama").value.trim();if(!nama){msg(m,"Isi nama anggota.",true);return}
 try{await getAnggota();const no=Math.max(0,...ANG.map(a=>a.no||0))+1,id="A"+String(no).padStart(3,"0"),inst=$("#a_inst").value.trim();
  const data={no,nama,instansi:inst,npa:$("#a_npa").value.trim(),hp:$("#a_hp").value.trim(),tglDaftar:today(),status:"aktif"};
  const b=writeBatch(db);b.set(doc(db,"anggota",id),data);b.set(doc(db,"publik_anggota",id),{no,nama,instansi:inst,pokok:0,wajib:0,sukarela:0,total:0});logOp(b,{aksi:"tambah",kol:"anggota",id,sesudah:data,detail:nama});await b.commit();
  ["a_nama","a_inst","a_npa","a_hp"].forEach(i=>$("#"+i).value="");await getAnggota(true);await syncRingkasan();await loadAnggota();msg(m,`Anggota nomor ${no} ditambahkan. Lanjut catat simpanan pokoknya di menu Simpanan.`)}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};

// ---------- Simpanan ----------
async function loadSimpanan(){const [ang,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);SIM=s.docs.map(d=>({id:d.id,...d.data()}));
 $("#s_form").classList.toggle("hide",!(can("simpanan","ubah")&&can("bku","ubah")));
 const opt=ang.map(a=>`<option value="${esc(a.id)}">${a.no}. ${esc(a.nama)}</option>`).join("");$("#s_ang").innerHTML=opt;$("#s_fa").innerHTML=`<option value="">Semua anggota</option>`+opt;
 if(!$("#s_tgl").value)$("#s_tgl").value=today();if(!$("#s_per").value)$("#s_per").value=today().slice(0,7);renderSim()}
function renderSim(){const q=$("#s_cari").value.toLowerCase(),t={};SIM.forEach(x=>{(t[x.anggotaId]??={pokok:0,wajib:0,sukarela:0})[x.jenis]+=x.jumlah});
 const r=ANG.filter(a=>((a.nama||"")+(a.instansi||"")).toLowerCase().includes(q)).map(a=>{const x=t[a.id]||{pokok:0,wajib:0,sukarela:0};return {a,...x,tot:x.pokok+x.wajib+x.sukarela}});
 $("#sb").innerHTML=r.map(x=>`<tr><td class="n">${x.a.no}</td><td>${esc(x.a.nama)}</td><td>${esc(x.a.instansi)}</td><td class="n">${rp(x.pokok)}</td><td class="n">${rp(x.wajib)}</td><td class="n">${rp(x.sukarela)}</td><td class="n">${rp(x.tot)}</td></tr>`).join("");
 const s=k=>r.reduce((a,x)=>a+x[k],0);$("#sf").innerHTML=`<tr><td></td><td colspan="2">Jumlah (${r.length} anggota)</td><td class="n">${rp(s("pokok"))}</td><td class="n">${rp(s("wajib"))}</td><td class="n">${rp(s("sukarela"))}</td><td class="n">${rp(s("tot"))}</td></tr>`;
 $("#s_xls").onclick=()=>unduh(r.map(x=>({No:x.a.no,Nama:x.a.nama,Instansi:x.a.instansi,Pokok:x.pokok,Wajib:x.wajib,Sukarela:x.sukarela,Total:x.tot})),"simpanan-anggota.xlsx");renderSetoran()}
function renderSetoran(){const f=$("#s_fa").value,e=can("simpanan","ubah")&&can("bku","ubah"),nm=Object.fromEntries(ANG.map(a=>[a.id,a.nama]));
 const r=SIM.filter(x=>!f||x.anggotaId===f).sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")).slice(0,300);
 $("#sh").innerHTML=r.map(x=>`<tr><td>${esc(x.tanggal||"")}</td><td>${esc(nm[x.anggotaId]||x.anggotaId)}</td><td>${esc(x.jenis)}</td><td>${esc(x.periode||"")}</td><td class="n">${rp(x.jumlah)}</td><td>${e?btn("ubah",x.id,"Ubah")+" "+btn("hapus",x.id,"Hapus"):""} ${rwBtn("simpanan",x.id)}</td></tr>`).join("")||`<tr><td colspan="6">Belum ada setoran.</td></tr>`}
$("#s_cari").oninput=renderSim;$("#s_fa").onchange=renderSetoran;
$("#s_jenis").onchange=()=>{const j=$("#s_jenis").value;$("#s_perwrap").classList.toggle("hide",j!=="wajib");$("#s_jml").value=j==="pokok"?500000:j==="wajib"?50000:""};
const cekSukarela=j=>j>15000000?"Simpanan sukarela maksimal Rp15.000.000 per anggota (AD Pasal 34).":"";
$("#s_save").onclick=async()=>{const m=$("#s_msg"),id=$("#s_ang").value,jenis=$("#s_jenis").value,jml=+$("#s_jml").value,tgl=$("#s_tgl").value,per=$("#s_per").value;
 if(!id||!(jml>0)||!tgl){msg(m,"Lengkapi anggota, jumlah, dan tanggal.",true);return}
 if(jenis==="wajib"&&!per){msg(m,"Pilih bulan untuk simpanan wajib.",true);return}
 if(jenis==="sukarela"&&cekSukarela(jml)){msg(m,cekSukarela(jml),true);return}
 const sid=jenis==="pokok"?id+"_pokok":jenis==="sukarela"?id+"_sukarela":`${id}_wajib_${per.slice(0,4)}_${per.slice(5,7)}`;
 if(SIM.some(x=>x.id===sid)){msg(m,"Setoran jenis dan bulan ini untuk anggota tersebut sudah tercatat. Gunakan tombol Ubah pada Daftar setoran.",true);return}
 $("#s_save").disabled=true;
 try{const a=ANG.find(x=>x.id===id),data={anggotaId:id,jenis,jumlah:jml,tanggal:tgl};if(jenis==="wajib")data.periode=per;
  const b=writeBatch(db);b.set(doc(db,"simpanan",sid),data);
  b.set(doc(db,"bku","sim_"+sid),{tanggal:tgl,kode:"SIM",keterangan:`Simpanan ${jenis}${jenis==="wajib"?" "+per:""} an: ${a.nama}`,kredit:jml,debit:0,ref:sid,anggotaId:id,jenis,urut:Date.now()});
  logOp(b,{aksi:"tambah",kol:"simpanan",id:sid,sesudah:data,detail:`${a.nama}: ${jenis}`});await b.commit();await syncPublik(id);
  const s=await getDocs(collection(db,"simpanan"));SIM=s.docs.map(d=>({id:d.id,...d.data()}));renderSim();msg(m,`Setoran ${jenis} ${a.nama} sebesar ${rp(jml)} tersimpan dan masuk ke buku kas.`)}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}
 $("#s_save").disabled=false};
aksi("#sh",{rw,
 ubah:async id=>{const x=SIM.find(y=>y.id===id),a=ANG.find(y=>y.id===x.anggotaId),kunci=!!x.bkuId;
  const v=await tanya({judul:`Ubah setoran ${x.jenis} · ${a?.nama||""}`,fields:[{type:"note",label:kunci?"Setoran ini berasal dari konfirmasi BKU. Jumlah dikunci; ubah lewat menu Buku kas (Batalkan konfirmasi, lalu konfirmasi ulang).":"Jenis dan bulan tidak bisa diubah. Jika salah, hapus setoran lalu catat ulang."},{k:"tanggal",label:"Tanggal setor",type:"date",req:1},{k:"jumlah",label:"Jumlah (Rp)",type:"number",req:1,dis:kunci}],nilai:{tanggal:x.tanggal,jumlah:x.jumlah},validasi:o=>!(o.jumlah>0)?"Jumlah harus lebih dari 0.":x.jenis==="sukarela"?cekSukarela(o.jumlah):""});if(!v)return;
  const sebelum={tanggal:x.tanggal||"",jumlah:x.jumlah},sesudah={tanggal:v.tanggal,jumlah:v.jumlah};if(JSON.stringify(sebelum)===JSON.stringify(sesudah))return alert("Tidak ada yang diubah.");
  const b=writeBatch(db);b.set(doc(db,"simpanan",id),sesudah,{merge:true});
  const bk=await getDoc(doc(db,"bku","sim_"+id));if(bk.exists())b.set(doc(db,"bku","sim_"+id),{tanggal:v.tanggal,kredit:v.jumlah},{merge:true});
  logOp(b,{aksi:"ubah",kol:"simpanan",id,sebelum,sesudah,alasan:v.alasan,detail:a?.nama});await b.commit();await syncPublik(x.anggotaId);await loadSimpanan()},
 hapus:async id=>{const x=SIM.find(y=>y.id===id),a=ANG.find(y=>y.id===x.anggotaId);
  const v=await tanya({judul:"Hapus setoran",tombol:"Hapus",fields:[{type:"note",label:`Hapus setoran ${x.jenis}${x.periode?" "+x.periode:""} ${rp(x.jumlah)} milik ${a?.nama||""}? Catatan terkait di buku kas ikut dihapus.`}]});if(!v)return;
  const b=writeBatch(db);b.delete(doc(db,"simpanan",id));b.delete(doc(db,"bku","sim_"+id));if(x.bkuId)b.set(doc(db,"bku",x.bkuId),{konfirmasi:deleteField()},{merge:true});
  const {id:_i,...snap}=x;logOp(b,{aksi:"hapus",kol:"simpanan",id,sebelum:snap,alasan:v.alasan,detail:a?.nama});await b.commit();await syncPublik(x.anggotaId);await loadSimpanan()}});

// ---------- Buku kas ----------
$("#b_kode").onchange=()=>$("#b_note").classList.toggle("hide",$("#b_kode").value!=="SIM");
const statusSim=x=>{if(x.kode!=="SIM")return "";if(x.ref)return "✓ Terkonfirmasi (Simpanan)";if(x.konfirmasi){const k=x.konfirmasi,n=ANG.find(a=>a.id===k.anggotaId)?.nama||"",p=[k.pokok&&"Pokok",k.wajib&&"Wajib",k.sukarela&&"Sukarela"].filter(Boolean).join("+");return `✓ ${n} · ${p}`}return "Perlu konfirmasi"};
async function loadBku(){const [s]=await Promise.all([getDocs(collection(db,"bku")),getAnggota()]);BKU=s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(a.tanggal||"").localeCompare(b.tanggal||"")||(a.urut||0)-(b.urut||0));
 $("#b_form").classList.toggle("hide",!can("bku","ubah"));if(!$("#b_tgl").value)$("#b_tgl").value=today();
 const kodes=[...new Set([...KODE,...BKU.map(x=>x.kode).filter(Boolean)])],cur=$("#b_fk").value;
 $("#b_fk").innerHTML=`<option value="">Semua kode</option>`+kodes.map(k=>`<option ${k===cur?"selected":""}>${esc(k)}</option>`).join("");renderBku()}
function renderBku(){const f=$("#b_fk").value,e=can("bku","ubah");let sal=0,m=0,k=0;const out=[];
 const pend=BKU.filter(x=>x.kode==="SIM"&&!x.ref&&!x.konfirmasi).length;
 $("#b_warn").innerHTML=pend?`<div class="warn"><span><b>${pend} transaksi SIM perlu konfirmasi</b>: tentukan anggota dan pembagiannya ke pokok, wajib, atau sukarela.</span></div>`:"";
 const rows=BKU.map((x,i)=>{sal+=(x.kredit||0)-(x.debit||0);return {x,i,sal}}).filter(r=>!f||r.x.kode===f);
 $("#bb").innerHTML=rows.map(({x,i,sal})=>{m+=x.kredit||0;k+=x.debit||0;const st=statusSim(x);out.push({No:i+1,Tanggal:x.tanggal,Kode:x.kode||"",Keterangan:x.keterangan,"Masuk (kredit)":x.kredit||0,"Keluar (debit)":x.debit||0,Saldo:sal,Status:st});
  const a=[];if(e){a.push(btn("ubah",x.id,"Ubah"));if(x.kode==="SIM"&&!x.ref){a.push(x.konfirmasi?btn("batal",x.id,"Batalkan konfirmasi"):btn("konf",x.id,"Konfirmasi"))}a.push(btn("hapus",x.id,"Hapus"))}a.push(rwBtn("bku",x.id));
  return `<tr><td class="n">${i+1}</td><td>${esc(x.tanggal)}</td><td>${esc(x.kode||"")}</td><td>${esc(x.keterangan)}</td><td class="n">${x.kredit?rp(x.kredit):""}</td><td class="n">${x.debit?rp(x.debit):""}</td><td class="n">${rp(sal)}</td><td class="${st==="Perlu konfirmasi"?"bad":""}">${esc(st)}</td><td>${a.join(" ")}</td></tr>`}).join("");
 $("#bf").innerHTML=`<tr><td colspan="4">Jumlah${f?" kode "+esc(f):""}</td><td class="n">${rp(m)}</td><td class="n">${rp(k)}</td><td class="n">${rp(sal)}</td><td colspan="2">saldo akhir seluruh kas</td></tr>`;
 $("#b_xls").onclick=()=>unduh(out,"buku-kas-umum.xlsx")}
$("#b_fk").onchange=renderBku;
$("#b_save").onclick=async()=>{const m=$("#b_msg"),jml=+$("#b_jml").value,ket=$("#b_ket").value.trim(),tgl=$("#b_tgl").value;
 if(!(jml>0)||!ket||!tgl){msg(m,"Lengkapi tanggal, keterangan, dan jumlah.",true);return}
 try{const masuk=$("#b_arah").value==="masuk",data={tanggal:tgl,kode:$("#b_kode").value,keterangan:ket,kredit:masuk?jml:0,debit:masuk?0:jml,urut:Date.now()},ref=doc(collection(db,"bku")),b=writeBatch(db);
  b.set(ref,data);logOp(b,{aksi:"tambah",kol:"bku",id:ref.id,sesudah:data,detail:ket});await b.commit();$("#b_jml").value="";$("#b_ket").value="";await loadBku();msg(m,"Transaksi tersimpan.")}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};
aksi("#bb",{rw,
 ubah:async id=>{const x=BKU.find(y=>y.id===id),kunci=!!(x.ref||x.konfirmasi?.mode==="catat");
  const ekstra=(!x.kode||!KODE.includes(x.kode))?[[x.kode||"",x.kode||"(tanpa kode)"]]:[];
  const v=await tanya({judul:"Ubah transaksi kas",fields:[{type:"note",label:kunci?"Transaksi ini terhubung ke data simpanan: kode, arah, dan jumlah dikunci. Ubah lewat menu Simpanan atau batalkan konfirmasi.":"Ubah isian yang keliru."},{k:"tanggal",label:"Tanggal",type:"date",req:1},{k:"kode",label:"Kode",type:"select",opt:[...ekstra,...KODE.map(c=>[c,c])],dis:kunci},{k:"arah",label:"Arah",type:"select",opt:[["masuk","Uang masuk (kredit)"],["keluar","Uang keluar (debit)"]],dis:kunci},{k:"jumlah",label:"Jumlah (Rp)",type:"number",req:1,dis:kunci},{k:"keterangan",label:"Keterangan",req:1}],
   nilai:{tanggal:x.tanggal,kode:x.kode||"",arah:x.kredit?"masuk":"keluar",jumlah:x.kredit||x.debit||0,keterangan:x.keterangan},validasi:o=>o.jumlah>0?"":"Jumlah harus lebih dari 0."});if(!v)return;
  const sebelum={tanggal:x.tanggal||"",kode:x.kode||"",keterangan:x.keterangan||"",kredit:x.kredit||0,debit:x.debit||0},sesudah={tanggal:v.tanggal,kode:v.kode,keterangan:v.keterangan,kredit:v.arah==="masuk"?v.jumlah:0,debit:v.arah==="keluar"?v.jumlah:0};
  if(JSON.stringify(sebelum)===JSON.stringify(sesudah))return alert("Tidak ada yang diubah.");
  const upd={...sesudah};if(x.konfirmasi&&sesudah.kode!=="SIM")upd.konfirmasi=deleteField();
  const b=writeBatch(db);b.set(doc(db,"bku",id),upd,{merge:true});logOp(b,{aksi:"ubah",kol:"bku",id,sebelum,sesudah,alasan:v.alasan,detail:v.keterangan});await b.commit();await loadBku()},
 hapus:async id=>{const x=BKU.find(y=>y.id===id);
  if(x.ref||x.konfirmasi?.mode==="catat")return alert("Transaksi ini terhubung ke data simpanan. Hapus lewat menu Simpanan, atau gunakan Batalkan konfirmasi terlebih dahulu.");
  const v=await tanya({judul:"Hapus transaksi kas",tombol:"Hapus",fields:[{type:"note",label:`Hapus transaksi "${x.keterangan}" (${rp(x.kredit||x.debit)})?`}]});if(!v)return;
  const {id:_i,...snap}=x;const b=writeBatch(db);b.delete(doc(db,"bku",id));logOp(b,{aksi:"hapus",kol:"bku",id,sebelum:snap,alasan:v.alasan,detail:x.keterangan});await b.commit();await loadBku()},
 konf:async id=>{const x=BKU.find(y=>y.id===id);await getAnggota();
  const nm=s=>(s||"").toLowerCase().replace(/[^a-z ]/g,"").trim(),k=nm(x.keterangan),tebak=ANG.find(a=>{const n=nm(a.nama);return n&&k&&(n===k||n.startsWith(k)||k.startsWith(n))});
  const v=await tanya({judul:"Konfirmasi simpanan (kode SIM)",fields:[{type:"note",label:`Setoran ${rp(x.kredit)} tanggal ${x.tanggal}. Bagi ke pokok, wajib, dan sukarela. Totalnya harus sama dengan jumlah setoran.`},{k:"anggotaId",label:"Anggota",type:"select",opt:ANG.map(a=>[a.id,`${a.no}. ${a.nama}`]),req:1},{k:"pokok",label:"Simpanan pokok (Rp)",type:"number"},{k:"wajib",label:"Simpanan wajib (Rp)",type:"number"},{k:"periode",label:"Bulan simpanan wajib",type:"month"},{k:"sukarela",label:"Simpanan sukarela (Rp)",type:"number"},{k:"mode",label:"Cara menyimpan",type:"select",opt:[["tandai","Tandai saja (data simpanan anggota sudah ada, mis. dari impor)"],["catat","Catat juga ke simpanan anggota (menimpa data yang jenis/bulannya sama)"]]}],
   nilai:{anggotaId:tebak?.id||ANG[0]?.id,pokok:0,wajib:0,sukarela:0,periode:(x.tanggal||"").slice(0,7),mode:String(x.id).startsWith("B")?"tandai":"catat"},
   validasi:o=>o.pokok+o.wajib+o.sukarela!==x.kredit?`Pokok + wajib + sukarela (${rp(o.pokok+o.wajib+o.sukarela)}) harus sama dengan ${rp(x.kredit)}.`:o.wajib>0&&!o.periode?"Pilih bulan simpanan wajib.":cekSukarela(o.sukarela)||(o.mode==="catat"&&!can("simpanan","ubah")?"Cara 'catat' membutuhkan hak Ubah pada Simpanan.":"")});if(!v)return;
  const b=writeBatch(db),refs=[],tg=x.tanggal;
  if(v.mode==="catat"){const a=v.anggotaId;
   if(v.pokok){refs.push(a+"_pokok");b.set(doc(db,"simpanan",a+"_pokok"),{anggotaId:a,jenis:"pokok",jumlah:v.pokok,tanggal:tg,bkuId:id})}
   if(v.wajib){const s=`${a}_wajib_${v.periode.slice(0,4)}_${v.periode.slice(5,7)}`;refs.push(s);b.set(doc(db,"simpanan",s),{anggotaId:a,jenis:"wajib",jumlah:v.wajib,tanggal:tg,periode:v.periode,bkuId:id})}
   if(v.sukarela){refs.push(a+"_sukarela");b.set(doc(db,"simpanan",a+"_sukarela"),{anggotaId:a,jenis:"sukarela",jumlah:v.sukarela,tanggal:tg,bkuId:id})}}
  const konf={anggotaId:v.anggotaId,pokok:v.pokok,wajib:v.wajib,sukarela:v.sukarela,periode:v.periode||"",mode:v.mode,refs};
  b.set(doc(db,"bku",id),{konfirmasi:konf},{merge:true});logOp(b,{aksi:"konfirmasi",kol:"bku",id,sesudah:konf,alasan:v.alasan,detail:x.keterangan});await b.commit();
  if(v.mode==="catat")await syncPublik(v.anggotaId);await loadBku()},
 batal:async id=>{const x=BKU.find(y=>y.id===id),k=x.konfirmasi;
  const v=await tanya({judul:"Batalkan konfirmasi",tombol:"Batalkan",fields:[{type:"note",label:k.mode==="catat"?"Data simpanan yang dibuat dari konfirmasi ini ikut dihapus.":"Hanya tanda konfirmasi yang dihapus; data simpanan tidak berubah."}]});if(!v)return;
  const b=writeBatch(db);b.set(doc(db,"bku",id),{konfirmasi:deleteField()},{merge:true});(k.refs||[]).forEach(s=>b.delete(doc(db,"simpanan",s)));
  logOp(b,{aksi:"batal_konfirmasi",kol:"bku",id,sebelum:k,alasan:v.alasan,detail:x.keterangan});await b.commit();if(k.mode==="catat")await syncPublik(k.anggotaId);await loadBku()}});

// ---------- Surat ----------
const ROMAWI=["","I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII"];
async function loadSurat(){const s=await getDocs(collection(db,"surat"));SURAT=s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")||(b.urut||0)-(a.urut||0));
 $("#r_form").classList.toggle("hide",!can("surat","ubah"));if(!$("#r_tgl").value)$("#r_tgl").value=today();nomorOtomatis();renderSurat()}
function nomorOtomatis(){const k=$("#r_jenis").value==="keluar";$("#r_pl").textContent=k?"Tujuan":"Asal surat";
 if(!k){$("#r_nomor").placeholder="Nomor dari pengirim";return}
 const tg=$("#r_tgl").value||today(),th=tg.slice(0,4),n=SURAT.filter(x=>x.jenis==="keluar"&&x.tahun===th).length+1;
 $("#r_nomor").value=`${String(n).padStart(3,"0")}/KOP.PGRI-BRU/${ROMAWI[+tg.slice(5,7)]}/${th}`}
$("#r_jenis").onchange=()=>{$("#r_nomor").value="";nomorOtomatis()};$("#r_tgl").onchange=()=>{if($("#r_jenis").value==="keluar")nomorOtomatis()};
function renderSurat(){const f=$("#r_f").value,e=can("surat","ubah");$("#rb").innerHTML=SURAT.filter(x=>!f||x.jenis===f).map(x=>`<tr><td>${esc(x.tanggal)}</td><td>${x.jenis==="keluar"?"Keluar":"Masuk"}</td><td>${esc(x.nomor)}</td><td>${esc(x.pihak)}</td><td>${esc(x.perihal)}</td><td>${/^https?:\/\//.test(x.tautan||"")?`<a href="${esc(x.tautan)}" target="_blank" rel="noopener">Buka</a>`:""}</td><td>${e?btn("ubah",x.id,"Ubah")+" "+btn("hapus",x.id,"Hapus"):""} ${rwBtn("surat",x.id)}</td></tr>`).join("")||`<tr><td colspan="7">Belum ada surat.</td></tr>`}
$("#r_f").onchange=renderSurat;
$("#r_save").onclick=async()=>{const m=$("#r_msg"),tgl=$("#r_tgl").value,hal=$("#r_hal").value.trim(),pihak=$("#r_pihak").value.trim(),nomor=$("#r_nomor").value.trim();
 if(!tgl||!hal||!pihak||!nomor){msg(m,"Lengkapi tanggal, nomor, asal/tujuan, dan perihal.",true);return}
 try{const data={jenis:$("#r_jenis").value,tanggal:tgl,tahun:tgl.slice(0,4),nomor,pihak,perihal:hal,tautan:$("#r_url").value.trim(),urut:Date.now()},ref=doc(collection(db,"surat")),b=writeBatch(db);
  b.set(ref,data);logOp(b,{aksi:"tambah",kol:"surat",id:ref.id,sesudah:data,detail:`${nomor} - ${hal}`});await b.commit();
  ["r_hal","r_pihak","r_url"].forEach(i=>$("#"+i).value="");await loadSurat();msg(m,"Surat tersimpan.")}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};
aksi("#rb",{rw,
 ubah:async id=>{const x=SURAT.find(y=>y.id===id);
  const v=await tanya({judul:"Ubah surat",fields:[{k:"nomor",label:"Nomor surat",req:1},{k:"tanggal",label:"Tanggal",type:"date",req:1},{k:"pihak",label:x.jenis==="keluar"?"Tujuan":"Asal surat",req:1},{k:"perihal",label:"Perihal",req:1},{k:"tautan",label:"Tautan berkas",type:"url"}],nilai:x});if(!v)return;
  const sesudah={nomor:v.nomor,tanggal:v.tanggal,tahun:v.tanggal.slice(0,4),pihak:v.pihak,perihal:v.perihal,tautan:v.tautan},sebelum={nomor:x.nomor||"",tanggal:x.tanggal||"",tahun:x.tahun||"",pihak:x.pihak||"",perihal:x.perihal||"",tautan:x.tautan||""};
  if(JSON.stringify(sebelum)===JSON.stringify(sesudah))return alert("Tidak ada yang diubah.");
  const b=writeBatch(db);b.set(doc(db,"surat",id),sesudah,{merge:true});logOp(b,{aksi:"ubah",kol:"surat",id,sebelum,sesudah,alasan:v.alasan,detail:v.nomor});await b.commit();await loadSurat()},
 hapus:async id=>{const x=SURAT.find(y=>y.id===id);
  const v=await tanya({judul:"Hapus surat",tombol:"Hapus",fields:[{type:"note",label:`Hapus surat nomor ${x.nomor}?`}]});if(!v)return;
  const {id:_i,...snap}=x;const b=writeBatch(db);b.delete(doc(db,"surat",id));logOp(b,{aksi:"hapus",kol:"surat",id,sebelum:snap,alasan:v.alasan,detail:x.nomor});await b.commit();await loadSurat()}});

// ---------- SHU ----------
const SHU_PCT={cadangan:40,pendidikan:5,pengurus:10,anggota:20,transaksi:15,sukarela:10};
const PENERIMA=["Marsono (Ketua)","Ika Novita Mulia (Sekretaris)","Eli Sutrisnaniati (Bendahara)","M. Romsyah (Koordinator Unit Usaha)","Hj. Mardiatul Idalisah (Ketua Pengawas)","Mustaring (Anggota Pengawas)","Dr. H. Saleh (Anggota Pengawas)"];
async function loadShu(){const [ang,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);
 const suk={};s.forEach(d=>{const x=d.data();if(x.jenis==="sukarela")suk[x.anggotaId]=(suk[x.anggotaId]||0)+x.jumlah});
 SHUROWS=ang.filter(a=>a.status!=="berhenti").map(a=>({id:a.id,nama:a.nama,suk:suk[a.id]||0,tx:0}));
 if(!$("#h_thn").value)$("#h_thn").value=new Date().getFullYear();if(!$("#h_pen").value)$("#h_pen").value=PENERIMA.join("\n");
 $("#h_save").classList.toggle("hide",!can("shu","ubah"));await muatShu();renderShu()}
async function muatShu(){$("#h_msg").classList.add("hide");try{const d=await getDoc(doc(db,"shu",String($("#h_thn").value)));
 if(d.exists()){const x=d.data(),t=Object.fromEntries((x.hasil||[]).map(r=>[r.id,r.transaksi]));$("#h_laba").value=x.laba;$("#h_pen").value=(x.penerima||[]).join("\n");SHUROWS.forEach(r=>r.tx=t[r.id]||0);msg($("#h_msg"),`Memuat perhitungan yang tersimpan untuk tahun ${x.tahun}.`)}}catch(e){}}
$("#h_thn").onchange=async()=>{await muatShu();renderShu()};
function renderShu(){$("#hb").innerHTML=SHUROWS.map(r=>`<tr><td>${esc(r.nama)}</td><td class="n">${rp(r.suk)}</td><td class="n"><input type="number" min="0" class="tx" data-id="${esc(r.id)}" value="${r.tx||""}" style="width:150px;text-align:right" aria-label="Transaksi ${esc(r.nama)}"></td><td class="n" id="hs_${esc(r.id)}"></td><td class="n" id="ht_${esc(r.id)}"></td><td class="n" id="ha_${esc(r.id)}"></td><td class="n" id="hz_${esc(r.id)}"></td></tr>`).join("");hitungShu()}
$("#hb").oninput=e=>{if(!e.target.classList.contains("tx"))return;const r=SHUROWS.find(x=>x.id===e.target.dataset.id);if(r)r.tx=+e.target.value||0;hitungShu()};
$("#h_laba").oninput=()=>hitungShu();$("#h_pen").oninput=()=>hitungShu();
function hitungShu(){const L=+$("#h_laba").value||0,pen=$("#h_pen").value.split("\n").map(s=>s.trim()).filter(Boolean);
 const P=Object.fromEntries(Object.entries(SHU_PCT).map(([k,v])=>[k,Math.round(L*v/100)]));
 const tS=SHUROWS.reduce((s,r)=>s+r.suk,0),tT=SHUROWS.reduce((s,r)=>s+r.tx,0),n=SHUROWS.length;
 SHUROWS.forEach(r=>{r.hs=tS?P.sukarela*r.suk/tS:0;r.ht=tT?P.transaksi*r.tx/tT:0;r.ha=n?P.anggota/n:0;r.z=r.hs+r.ht+r.ha;
  for(const [p,k] of [["hs","hs"],["ht","ht"],["ha","ha"],["hz","z"]]){const c=$(`#${p}_${r.id}`);if(c)c.textContent=rp(r[k])}});
 const each=pen.length?P.pengurus/pen.length:0;
 $("#h_cards").innerHTML=[["Kas cadangan koperasi ("+SHU_PCT.cadangan+"%)",P.cadangan],["Dana pendidikan & sosial ("+SHU_PCT.pendidikan+"%)",P.pendidikan],["Pengurus & pengawas ("+SHU_PCT.pengurus+"%)",P.pengurus],["SHU anggota ("+SHU_PCT.anggota+"%)",P.anggota],["SHU transaksi anggota ("+SHU_PCT.transaksi+"%)",P.transaksi],["SHU simpanan sukarela ("+SHU_PCT.sukarela+"%)",P.sukarela],["Per penerima pengurus/pengawas",each]].map(([l,v])=>`<div class="stat"><b>${rp(v)}</b><span>${l}</span></div>`).join("");
 const w=[];if(L>0&&!tS)w.push("Belum ada simpanan sukarela, jadi bagian 10% belum bisa dibagikan.");if(L>0&&!tT)w.push("Isi kolom transaksi anggota agar bagian 15% bisa dibagikan.");if(L>0&&!pen.length)w.push("Isi penerima bagian pengurus & pengawas.");
 $("#h_note").innerHTML=w.map(t=>`<div class="msg err">${t}</div>`).join("");
 const s=k=>SHUROWS.reduce((a,r)=>a+r[k],0);$("#hf").innerHTML=`<tr><td>Jumlah</td><td class="n">${rp(tS)}</td><td class="n">${rp(tT)}</td><td class="n">${rp(s("hs"))}</td><td class="n">${rp(s("ht"))}</td><td class="n">${rp(s("ha"))}</td><td class="n">${rp(s("z"))}</td></tr>`;
 $("#h_xls").onclick=()=>unduh(SHUROWS.map(r=>({Anggota:r.nama,"Simpanan sukarela":r.suk,Transaksi:r.tx,"SHU sukarela":Math.round(r.hs),"SHU transaksi":Math.round(r.ht),"SHU anggota":Math.round(r.ha),Total:Math.round(r.z)})),`SHU-${$("#h_thn").value}.xlsx`)}
$("#h_save").onclick=async()=>{const m=$("#h_msg"),L=+$("#h_laba").value,thn=String($("#h_thn").value),pen=$("#h_pen").value.split("\n").map(s=>s.trim()).filter(Boolean);
 if(!(L>0)){msg(m,"Isi laba bersih terlebih dahulu.",true);return}
 try{const lama=await getDoc(doc(db,"shu",thn)),data={tahun:+thn,laba:L,persen:SHU_PCT,penerima:pen,perPenerima:pen.length?Math.round(L*SHU_PCT.pengurus/100/pen.length):0,
  hasil:SHUROWS.map(r=>({id:r.id,nama:r.nama,sukarela:r.suk,transaksi:r.tx,shuSukarela:Math.round(r.hs),shuTransaksi:Math.round(r.ht),shuAnggota:Math.round(r.ha),total:Math.round(r.z)})),disimpan:new Date().toISOString(),oleh:me.nama};
  const b=writeBatch(db);b.set(doc(db,"shu",thn),data);logOp(b,{aksi:lama.exists()?"ubah":"tambah",kol:"shu",id:thn,sebelum:lama.exists()?{laba:lama.data().laba}:null,sesudah:{laba:L},detail:`SHU tahun ${thn}`,alasan:lama.exists()?"Perhitungan ulang":""});await b.commit();
  msg(m,`Perhitungan SHU tahun ${thn} tersimpan.`)}catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};

// ---------- Impor Excel ----------
$("#imp").onclick=async()=>{const m=$("#imsg"),file=$("#f").files[0];if(!file){msg(m,"Pilih file Excel dulu.",true);return}
 $("#imp").disabled=true;msg(m,"Membaca file…");
 try{const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true}),wsS=wb.Sheets["SIMPANAN ANGGOTA"],wsB=wb.Sheets["BKU"];
  if(!wsS&&!wsB)throw new Error('Sheet "SIMPANAN ANGGOTA" atau "BKU" tidak ditemukan');
  const ops=[],TAHUN=2026;let nA=0,nB=0;
  if(wsS){const rows=XLSX.utils.sheet_to_json(wsS,{header:1,defval:null}).filter(r=>typeof r[0]==="number"&&typeof r[2]==="string"&&r[2].trim());const tot={pokok:0,wajib:0,sukarela:0};
   for(const r of rows){const no=r[0],id="A"+String(no).padStart(3,"0"),tgl=r[1] instanceof Date?ymd(r[1]):"",pokok=+r[4]||0,suk=+r[5]||0;let wajib=0;
    ops.push([doc(db,"anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),tglDaftar:tgl}]);
    if(pokok)ops.push([doc(db,"simpanan",id+"_pokok"),{anggotaId:id,jenis:"pokok",jumlah:pokok,tanggal:tgl}]);
    if(suk)ops.push([doc(db,"simpanan",id+"_sukarela"),{anggotaId:id,jenis:"sukarela",jumlah:suk,tanggal:tgl}]);
    for(let b=0;b<12;b++){const v=+r[6+b]||0;if(v){wajib+=v;ops.push([doc(db,"simpanan",`${id}_wajib_${TAHUN}_${String(b+1).padStart(2,"0")}`),{anggotaId:id,jenis:"wajib",jumlah:v,periode:`${TAHUN}-${String(b+1).padStart(2,"0")}`}])}}
    tot.pokok+=pokok;tot.wajib+=wajib;tot.sukarela+=suk;
    ops.push([doc(db,"publik_anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),pokok,wajib,sukarela:suk,total:pokok+wajib+suk}])}
   ops.push([doc(db,"publik_ringkasan","utama"),{jumlahAnggota:rows.length,...tot,diperbarui:new Date().toISOString()}]);nA=rows.length}
  if(wsB&&can("bku","ubah")){const rows=XLSX.utils.sheet_to_json(wsB,{header:1,defval:null}).filter(r=>typeof r[0]==="number"&&r[1] instanceof Date&&typeof r[3]==="string");
   for(const r of rows)ops.push([doc(db,"bku","B"+String(r[0]).padStart(4,"0")),{tanggal:ymd(r[1]),kode:String(r[2]||"").trim().toUpperCase(),keterangan:r[3].trim(),kredit:+r[4]||0,debit:+r[5]||0,urut:r[0]}]);nB=rows.length}
  for(let i=0;i<ops.length;i+=400){const b=writeBatch(db);ops.slice(i,i+400).forEach(([ref,d])=>b.set(ref,d,{merge:true}));await b.commit();msg(m,`Menyimpan… ${Math.min(i+400,ops.length)}/${ops.length}`)}
  ANG=[];msg(m,`Selesai: ${nA} anggota dan ${nB} baris buku kas diimpor (${ops.length} catatan). Jika ada kode SIM di BKU, konfirmasi pembagiannya di menu Buku kas.`);await catat("impor_excel",`${nA} anggota, ${nB} baris BKU`)}
 catch(e){msg(m,"Impor gagal: "+e.message,true)}
 $("#imp").disabled=false};

// ---------- Log ----------
async function loadLog(){const s=await getDocs(query(collection(db,"log"),orderBy("waktu","desc"),limit(200)));const f=$("#l_f").value;
 $("#lb").innerHTML=s.docs.map(d=>d.data()).filter(x=>!f||x.koleksi===f).map(x=>`<tr><td>${esc(wkt(x))}</td><td>${esc(x.nama)}</td><td>${esc(x.aksi)}</td><td>${esc(x.koleksi||"")}${x.detail?" · "+esc(x.detail):""}</td><td>${esc(diffText(x))}</td><td>${esc(x.alasan||"")}</td></tr>`).join("")||`<tr><td colspan="6">Belum ada aktivitas.</td></tr>`}
$("#l_f").onchange=()=>loadLog();

// ---------- Pengguna (master) ----------
function buildIzin(sel){$("#izin").innerHTML=`<tr><th>Modul</th><th>Lihat</th><th>Ubah</th></tr>`+MOD.map(([m,l])=>`<tr><td>${l}</td>`+["lihat","ubah"].map(k=>`<td><input type="checkbox" data-m="${m}" data-k="${k}" aria-label="${l}: ${k}" ${(sel[k]||[]).includes(m)?"checked":""}></td>`).join("")+`</tr>`).join("")}
$("#izin").onchange=e=>{const c=e.target;if(!c.dataset.m)return;const row=c.closest("tr"),l=row.querySelector("[data-k=lihat]"),u=row.querySelector("[data-k=ubah]");
 if(c.dataset.k==="ubah"&&c.checked)l.checked=true;if(c.dataset.k==="lihat"&&!c.checked)u.checked=false};
const collectIzin=()=>{const r={lihat:[],ubah:[]};document.querySelectorAll("#izin input").forEach(c=>{if(c.checked)r[c.dataset.k].push(c.dataset.m)});return r};
function syncMaster(){$("#izin").classList.toggle("hide",$("#u_master").checked)}
$("#u_master").onchange=syncMaster;
function resetForm(){editing=null;$("#ftitle").textContent="Tambah akun";["u_nama","u_email","u_pw"].forEach(i=>$("#"+i).value="");$("#u_jab").value="lainnya";
 $("#u_email").disabled=false;$("#pwwrap").classList.remove("hide");$("#u_aktif").checked=true;$("#u_aktif").disabled=false;$("#u_master").checked=false;$("#u_master").disabled=false;$("#u_mail").classList.add("hide");$("#u_msg").classList.add("hide");buildIzin({});syncMaster()}
$("#u_reset").onclick=resetForm;
const ringkasIzin=u=>u.master?"Semua (master)":`Lihat: ${(u.lihat||[]).map(m=>LBL[m]||m).join(", ")||"-"} · Ubah: ${(u.ubah||[]).map(m=>LBL[m]||m).join(", ")||"-"}`;
async function loadUsers(){const s=await getDocs(collection(db,"admins"));USERS=s.docs.map(d=>({uid:d.id,...d.data()}));
 $("#ub").innerHTML=USERS.map(u=>`<tr><td>${esc(u.nama)}</td><td>${esc(u.jabatan||u.role||"")}</td><td>${esc(u.email||"")}</td><td>${u.aktif===false?"nonaktif":"aktif"}</td><td>${esc(ringkasIzin(u))}</td><td><button class="btn ghost sm" data-uid="${esc(u.uid)}">Ubah</button></td></tr>`).join("")}
$("#ub").onclick=e=>{const b=e.target.closest("button[data-uid]");if(!b)return;const u=USERS.find(x=>x.uid===b.dataset.uid);if(!u)return;
 editing=u;$("#ftitle").textContent="Ubah akun";$("#u_nama").value=u.nama||"";$("#u_jab").value=u.jabatan||u.role||"lainnya";$("#u_email").value=u.email||"";$("#u_email").disabled=true;
 $("#pwwrap").classList.add("hide");$("#u_aktif").checked=u.aktif!==false;$("#u_aktif").disabled=u.uid===me.uid;$("#u_master").checked=!!u.master;$("#u_master").disabled=u.uid===me.uid;buildIzin(u);syncMaster();
 $("#u_mail").classList.toggle("hide",!u.email);$("#u_msg").classList.add("hide");scrollTo({top:0,behavior:"smooth"})};
$("#u_mail").onclick=async()=>{try{await sendPasswordResetEmail(auth,editing.email);msg($("#u_msg"),"Email atur ulang kata sandi dikirim ke "+editing.email)}catch(e){msg($("#u_msg"),"Gagal mengirim: "+e.message,true)}};
$("#u_save").onclick=async()=>{const m=$("#u_msg"),nama=$("#u_nama").value.trim(),jab=$("#u_jab").value,email=$("#u_email").value.trim(),izin=collectIzin();
 if(!nama){msg(m,"Isi nama terlebih dahulu.",true);return}
 $("#u_save").disabled=true;
 try{
  if(editing){const data={nama,jabatan:jab},u=editing,sebelum={nama:u.nama||"",jabatan:u.jabatan||u.role||"",aktif:u.aktif!==false,master:!!u.master,lihat:u.lihat||[],ubah:u.ubah||[]};
   if(u.uid!==me.uid){data.aktif=$("#u_aktif").checked;data.master=$("#u_master").checked;data.lihat=izin.lihat;data.ubah=izin.ubah}
   const sesudah={...sebelum,...data},b=writeBatch(db);b.set(doc(db,"admins",u.uid),data,{merge:true});logOp(b,{aksi:"ubah",kol:"admins",id:u.uid,sebelum,sesudah,detail:nama});await b.commit();await loadUsers();msg(m,"Perubahan tersimpan.")}
  else{const pw=$("#u_pw").value;
   if(!email){msg(m,"Isi email terlebih dahulu.",true);$("#u_save").disabled=false;return}
   if(pw.length<12||!/[a-z]/.test(pw)||!/[A-Z]/.test(pw)||!/\d/.test(pw)){msg(m,"Kata sandi minimal 12 karakter dan memuat huruf besar, huruf kecil, dan angka.",true);$("#u_save").disabled=false;return}
   const sec=initializeApp(firebaseConfig,"sec"+Date.now());let uid;
   try{const c=await createUserWithEmailAndPassword(getAuth(sec),email,pw);uid=c.user.uid;await signOut(getAuth(sec))}finally{await deleteApp(sec)}
   const data={nama,jabatan:jab,email,aktif:$("#u_aktif").checked,master:$("#u_master").checked,lihat:izin.lihat,ubah:izin.ubah},b=writeBatch(db);
   b.set(doc(db,"admins",uid),data);logOp(b,{aksi:"tambah",kol:"admins",id:uid,sesudah:data,detail:`${nama} (${email})`});await b.commit();await loadUsers();resetForm();msg(m,`Akun ${email} dibuat. Sampaikan kata sandi awal secara pribadi dan minta diganti setelah masuk.`)}}
 catch(e){msg(m,e.code==="auth/email-already-in-use"?"Email ini sudah terdaftar.":"Gagal menyimpan: "+e.message,true)}
 $("#u_save").disabled=false};
