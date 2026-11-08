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
const MOD=[["anggota","Anggota"],["simpanan","Simpanan"],["bku","Buku kas (BKU)"],["surat","Surat masuk/keluar"],["shu","SHU"],["tutup","Tutup tahun buku"],["log","Log & riwayat"]];
const LBL=Object.fromEntries(MOD);
const KODE=["SIM","BIA","CAFE","KONVEKSI","PINJ","SHU","CAD"],TIPE={SIM:"simpanan",BIA:"umum",CAFE:"usaha",KONVEKSI:"usaha",PINJ:"piutang",SHU:"dist_shu",CAD:"memo"},KLBL={SIM:"simpanan",BIA:"biaya lain-lain",CAFE:"kas cafe",KONVEKSI:"kas konveksi",PINJ:"peminjaman dana (piutang)",SHU:"pembagian SHU",CAD:"dana cadangan (pencatatan)"};
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
const aksi=(sel,map)=>{$(sel).onclick=async e=>{const b=e.target.closest("button[data-a]");if(!b)return;b.disabled=true;try{await map[b.dataset.a](b.dataset.id,b)}catch(err){console.error(err);alert(err.code==="permission-denied"?"Ditolak. Kemungkinan tahun buku transaksi ini sudah ditutup, atau hak akses Anda tidak cukup.":"Gagal: "+err.message)}b.disabled=false}};
const rw=(id,b)=>riwayat(b.dataset.k,id);

// ---------- Navigasi ----------
const VIEWS=[
 ["ringkasan","Ringkasan",()=>can("anggota","lihat")&&can("simpanan","lihat"),()=>loadRingkasan()],
 ["anggota","Anggota",()=>can("anggota","lihat"),()=>loadAnggota()],
 ["simpanan","Simpanan",()=>can("simpanan","lihat")&&can("anggota","lihat"),()=>loadSimpanan()],
 ["tunggakan","Tunggakan",()=>can("simpanan","lihat")&&can("anggota","lihat"),()=>loadTunggakan()],
 ["bku","Buku kas",()=>can("bku","lihat"),()=>loadBku()],
 ["laporan","Laporan",()=>can("bku","lihat")&&can("simpanan","lihat")&&can("anggota","lihat"),()=>loadLaporan()],
 ["surat","Surat",()=>can("surat","lihat"),()=>loadSurat()],
 ["tutup","Tutup buku",()=>can("tutup","lihat")&&can("bku","lihat")&&can("simpanan","lihat"),()=>loadTutup()],
 ["shu","SHU",()=>can("shu","lihat")&&can("anggota","lihat")&&can("simpanan","lihat"),()=>loadShu()],
 ["impor","Impor Excel",()=>can("anggota","ubah")&&can("simpanan","ubah"),()=>{}],
 ["pengguna","Pengguna",()=>me.master,()=>{resetForm();loadUsers()}],
 ["log","Log",()=>can("log","lihat"),()=>loadLog()]];
async function show(t){VIEWS.forEach(v=>$("#"+v[0]).classList.toggle("hide",v[0]!==t));
 document.querySelectorAll("#tabs button").forEach(b=>b.setAttribute("aria-selected",b.dataset.t===t));
 document.querySelectorAll(".loaderr").forEach(e=>e.remove());
 try{await Promise.all([muatTB().catch(()=>{}),muatKode().catch(()=>{})]);await VIEWS.find(v=>v[0]===t)[3]()}catch(e){console.error(e);$("#"+t).insertAdjacentHTML("afterbegin",`<div class="msg err loaderr">Gagal memuat: ${esc(e.message)}</div>`)}}
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
const pubW=(a,L)=>({wajibBulan:[...new Set(L.filter(x=>x.jenis==="wajib"&&x.periode).map(x=>x.periode))].sort(),daftar:(a.tglDaftar||"").slice(0,7)});
async function syncPublik(id){const a=(await getAnggota()).find(x=>x.id===id);if(!a)return;
 if(a.status==="berhenti"){await deleteDoc(doc(db,"publik_anggota",id))}
 else{const q=await getDocs(query(collection(db,"simpanan"),where("anggotaId","==",id))),t={pokok:0,wajib:0,sukarela:0};q.forEach(d=>{const x=d.data();t[x.jenis]+=x.jumlah});
  await setDoc(doc(db,"publik_anggota",id),{no:a.no,nama:a.nama,instansi:a.instansi||"",...t,total:t.pokok+t.wajib+t.sukarela,...pubW(a,q.docs.map(d=>d.data()))},{merge:true})}
 await syncRingkasan()}

// ---------- Tunggakan (dipakai juga oleh Ringkasan) ----------
function hitungTunggakan(ang,sim,tahun,jt){const n=new Date();let due=TBD[tahun]?.dibagi?0:tahun<n.getFullYear()?12:tahun>n.getFullYear()?0:n.getMonth()+(n.getDate()>=jt?1:0);
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
 $("#ab").innerHTML=r.map(a=>`<tr><td class="n">${a.no}</td><td>${esc(a.nama)}${(!a.npa||!a.hp||!a.instansi||!a.tglDaftar)?` <span class="bad" title="Ada isian kosong: klik Ubah untuk melengkapi">⚠ belum lengkap</span>`:""}</td><td>${esc(a.instansi)}</td><td>${esc(a.npa||"")}</td><td>${esc(a.hp||"")}</td><td>${esc(a.tglDaftar||"")}</td><td>${esc(a.status||"aktif")}</td><td>${e?btn("ubah",a.id,"Ubah"):""} ${rwBtn("anggota",a.id)}</td></tr>`).join("")||`<tr><td colspan="8">Belum ada anggota. Tambahkan di atas atau gunakan Impor Excel.</td></tr>`}
aksi("#ab",{rw,ubah:async id=>{const a=ANG.find(x=>x.id===id);
 const v=await tanya({judul:`Ubah anggota No. ${a.no}`,fields:[{k:"nama",label:"Nama",req:1},{k:"instansi",label:"Instansi / sekolah"},{k:"npa",label:"NPA PGRI"},{k:"hp",label:"No HP"},{k:"alamat",label:"Alamat"},{k:"tglDaftar",label:"Tanggal daftar",type:"date"},{k:"status",label:"Status",type:"select",opt:[["aktif","Aktif"],["berhenti","Berhenti"]]}],nilai:{...a,status:a.status||"aktif"}});if(!v)return;
 const sesudah={nama:v.nama,instansi:v.instansi,npa:v.npa,hp:v.hp,alamat:v.alamat,tglDaftar:v.tglDaftar,status:v.status},sebelum={nama:a.nama||"",instansi:a.instansi||"",npa:a.npa||"",hp:a.hp||"",alamat:a.alamat||"",tglDaftar:a.tglDaftar||"",status:a.status||"aktif"};
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
 $("#sh").innerHTML=r.map(x=>`<tr><td>${esc(x.tanggal||"")}</td><td>${esc(nm[x.anggotaId]||x.anggotaId)}</td><td>${esc(x.tarik?"tarik sukarela":x.jenis)}</td><td>${esc(x.periode||"")}</td><td class="n">${rp(x.jumlah)}</td><td>${e?(x.tarik?"":btn("ubah",x.id,"Ubah")+" ")+btn("hapus",x.id,"Hapus"):""} ${rwBtn("simpanan",x.id)}</td></tr>`).join("")||`<tr><td colspan="6">Belum ada setoran.</td></tr>`}
$("#s_cari").oninput=renderSim;$("#s_fa").onchange=renderSetoran;
$("#s_jenis").onchange=()=>{const j=$("#s_jenis").value;$("#s_perwrap").classList.toggle("hide",j!=="wajib");$("#s_jml").value=j==="pokok"?500000:j==="wajib"?50000:""};
const cekSukarela=j=>j>15000000?"Simpanan sukarela maksimal Rp15.000.000 per anggota (AD Pasal 34).":"";
$("#s_save").onclick=async()=>{const m=$("#s_msg"),id=$("#s_ang").value,jenis=$("#s_jenis").value,jml=+$("#s_jml").value,tgl=$("#s_tgl").value,per=$("#s_per").value;
 if(!id||!(jml>0)||!tgl){msg(m,"Lengkapi anggota, jumlah, dan tanggal.",true);return}
 if(jenis==="wajib"&&!per){msg(m,"Pilih bulan untuk simpanan wajib.",true);return}
 if(tertutup(tgl)){msg(m,`Tahun buku ${tgl.slice(0,4)} sudah ditutup.`,true);return}
 if(jenis==="sukarela"||jenis==="tarik"){const sal=SIM.filter(x=>x.anggotaId===id&&x.jenis==="sukarela").reduce((t,x)=>t+x.jumlah,0),e=jenis==="tarik"?(jml>sal?`Penarikan melebihi saldo sukarela (${rp(sal)}).`:""):(sal+jml>15000000?"Simpanan sukarela maksimal Rp15.000.000 per anggota (AD Pasal 34).":"");if(e){msg(m,e,true);return}}
 const sid=jenis==="pokok"?id+"_pokok":jenis==="sukarela"||jenis==="tarik"?`${id}_${jenis}_${Date.now()}`:`${id}_wajib_${per.slice(0,4)}_${per.slice(5,7)}`;
 if(SIM.some(x=>x.id===sid)){msg(m,"Setoran jenis dan bulan ini untuk anggota tersebut sudah tercatat. Gunakan tombol Ubah pada Daftar setoran.",true);return}
 $("#s_save").disabled=true;
 try{const a=ANG.find(x=>x.id===id),tarik=jenis==="tarik",J=tarik?"sukarela":jenis,data={anggotaId:id,jenis:J,jumlah:tarik?-jml:jml,tanggal:tgl};if(tarik)data.tarik=true;if(jenis==="wajib")data.periode=per;
  const b=writeBatch(db);b.set(doc(db,"simpanan",sid),data);
  b.set(doc(db,"bku","sim_"+sid),{tanggal:tgl,kode:"SIM",keterangan:tarik?`Penarikan simpanan sukarela an: ${a.nama}`:`Simpanan ${jenis}${jenis==="wajib"?" "+per:""} an: ${a.nama}`,kredit:tarik?0:jml,debit:tarik?jml:0,ref:sid,anggotaId:id,jenis:J,urut:Date.now()});
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
 const rows=BKU.map((x,i)=>{sal+=(x.kredit||0)-(x.debit||0);return {x,i,sal}}).filter(r=>(!f||r.x.kode===f)&&(!$("#b_lk").checked||kurang(r.x).length));
 $("#bb").innerHTML=rows.map(({x,i,sal})=>{m+=x.kredit||0;k+=x.debit||0;const st=statusSim(x);out.push({No:i+1,Tanggal:x.tanggal,Kode:x.kode||"",Keterangan:x.keterangan,"Masuk (kredit)":x.kredit||0,"Keluar (debit)":x.debit||0,Saldo:sal,Status:st});
  const a=[];if(e){a.push(btn("ubah",x.id,"Ubah"));if(x.kode==="SIM"&&!x.ref){a.push(x.konfirmasi?btn("batal",x.id,"Batalkan konfirmasi"):btn("konf",x.id,"Konfirmasi"))}a.push(btn("hapus",x.id,"Hapus"))}a.push(rwBtn("bku",x.id));
  return `<tr><td class="n">${i+1}</td><td>${esc(x.tanggal)}</td><td>${esc(x.kode||"")}</td><td>${esc(x.keterangan)}<br><small class="dim">${esc(x.noBukti?"Bukti "+x.noBukti+" · ":"")}${esc(ALBL[akunDari(x)]||"")}${x.akun?"":" (otomatis)"}</small>${kurang(x).length?`<br><small class="bad">⚠ Lengkapi: ${kurang(x).join(", ")}</small>`:""}</td><td class="n">${x.kredit?rp(x.kredit):""}</td><td class="n">${x.debit?rp(x.debit):""}</td><td class="n">${rp(sal)}</td><td class="${st==="Perlu konfirmasi"?"bad":""}">${esc(st)}</td><td>${a.join(" ")}</td></tr>`}).join("");
 $("#bf").innerHTML=`<tr><td colspan="4">Jumlah${f?" kode "+esc(f):""}</td><td class="n">${rp(m)}</td><td class="n">${rp(k)}</td><td class="n">${rp(sal)}</td><td colspan="2">saldo akhir seluruh kas</td></tr>`;
 $("#b_xls").onclick=()=>unduh(out,"buku-kas-umum.xlsx")}
$("#b_fk").onchange=renderBku;$("#b_lk").onchange=renderBku;
$("#b_save").onclick=async()=>{const m=$("#b_msg"),jml=+$("#b_jml").value,ket=$("#b_ket").value.trim(),tgl=$("#b_tgl").value;
 if(!(jml>0)||!ket||!tgl){msg(m,"Lengkapi tanggal, keterangan, dan jumlah.",true);return}
 if(tertutup(tgl)){msg(m,`Tahun buku ${tgl.slice(0,4)} sudah ditutup.`,true);return}
 try{const masuk=$("#b_arah").value==="masuk",data={tanggal:tgl,kode:$("#b_kode").value,keterangan:ket,kredit:masuk?jml:0,debit:masuk?0:jml,noBukti:$("#b_bukti").value.trim(),pihak:$("#b_pihak").value.trim(),akun:$("#b_akun").value,urut:Date.now()},ref=doc(collection(db,"bku")),b=writeBatch(db);
  b.set(ref,data);logOp(b,{aksi:"tambah",kol:"bku",id:ref.id,sesudah:data,detail:ket});await b.commit();$("#b_jml").value="";$("#b_ket").value="";$("#b_bukti").value="";$("#b_pihak").value="";await loadBku();msg(m,"Transaksi tersimpan.")}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};
aksi("#bb",{rw,
 ubah:async id=>{const x=BKU.find(y=>y.id===id),kunci=!!(x.ref||x.konfirmasi?.mode==="catat");
  const ekstra=(!x.kode||!KODE.includes(x.kode))?[[x.kode||"",x.kode||"(tanpa kode)"]]:[];
  const v=await tanya({judul:"Ubah transaksi kas",fields:[{type:"note",label:kunci?"Transaksi ini terhubung ke data simpanan: kode, arah, jumlah, dan kategori dikunci (tanggal, keterangan, dan no. bukti tetap bisa dilengkapi). Ubah lewat menu Simpanan atau batalkan konfirmasi.":"Ubah isian yang keliru."},{k:"tanggal",label:"Tanggal",type:"date",req:1},{k:"kode",label:"Kode",type:"select",opt:[...ekstra,...KODE.map(c=>[c,c])],dis:kunci},{k:"arah",label:"Arah",type:"select",opt:[["masuk","Uang masuk (kredit)"],["keluar","Uang keluar (debit)"]],dis:kunci},{k:"jumlah",label:"Jumlah (Rp)",type:"number",req:1,dis:kunci},{k:"keterangan",label:"Keterangan",req:1},{k:"noBukti",label:"No. bukti / kuitansi"},{k:"pihak",label:"Peminjam / pihak (piutang, utang)"},{k:"akun",label:"Kategori laporan keuangan",type:"select",opt:AKUN.map(([v,l])=>[v,l]),dis:kunci}],
   nilai:{noBukti:x.noBukti||"",pihak:x.pihak||"",akun:x.akun||"",tanggal:x.tanggal,kode:x.kode||"",arah:x.kredit?"masuk":"keluar",jumlah:x.kredit||x.debit||0,keterangan:x.keterangan},validasi:o=>o.jumlah>0?"":"Jumlah harus lebih dari 0."});if(!v)return;
  const sebelum={tanggal:x.tanggal||"",kode:x.kode||"",keterangan:x.keterangan||"",kredit:x.kredit||0,debit:x.debit||0,noBukti:x.noBukti||"",pihak:x.pihak||"",akun:x.akun||""},sesudah={tanggal:v.tanggal,kode:v.kode,keterangan:v.keterangan,noBukti:v.noBukti,pihak:v.pihak,akun:kunci?(x.akun||""):v.akun,kredit:v.arah==="masuk"?v.jumlah:0,debit:v.arah==="keluar"?v.jumlah:0};
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
const thSHU=()=>+$("#h_thn").value,idn=n=>Math.round(n).toLocaleString("id-ID");
function bobotSukarela(th){const m={},out={};SIM.filter(x=>x.jenis==="sukarela").forEach(x=>(m[x.anggotaId]??=[]).push([x.tanggal||"",x.jumlah]));
 for(const [id,L] of Object.entries(m)){let w=0;for(let b=1;b<=12;b++){const e=`${th}-${String(b).padStart(2,"0")}-31`,sal=L.filter(([d])=>d<=e).reduce((t,[,j])=>t+j,0);if(sal>0)w+=sal}out[id]=w}return out}
function dasarShu(){const th=thSHU(),w=bobotSukarela(th),suk={},wj={},old=Object.fromEntries(SHUROWS.map(r=>[r.id,r.tx]));
 SIM.forEach(x=>{if(x.jenis==="sukarela")suk[x.anggotaId]=(suk[x.anggotaId]||0)+x.jumlah;if(x.jenis==="wajib"&&(x.periode||"").startsWith(th+"-"))wj[x.anggotaId]=(wj[x.anggotaId]||0)+x.jumlah});
 SHUROWS=ANG.filter(a=>a.status!=="berhenti").map(a=>({id:a.id,nama:a.nama,instansi:a.instansi||"",suk:suk[a.id]||0,w:w[a.id]||0,wj:wj[a.id]||0,tx:old[a.id]||0}))}
async function loadShu(){const [,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);SIM=s.docs.map(d=>({id:d.id,...d.data()}));
 if(!$("#h_thn").value)$("#h_thn").value=new Date().getFullYear();if(!$("#h_pen").value)$("#h_pen").value=PENERIMA.join("\n");
 $("#h_save").classList.toggle("hide",!can("shu","ubah"));$("#h_bagi").classList.toggle("hide",!(can("shu","ubah")&&can("bku","ubah")&&can("simpanan","ubah")));dasarShu();await muatShu();renderShu()}
async function muatShu(){$("#h_msg").classList.add("hide");try{const d=await getDoc(doc(db,"shu",String(thSHU())));
 if(d.exists()){const x=d.data(),h=Object.fromEntries((x.hasil||[]).map(r=>[r.id,r]));$("#h_laba").value=x.laba;$("#h_pen").value=(x.penerima||[]).join("\n");$("#h_cad").value=x.cadangan||"tetap";
  SHUROWS.forEach(r=>{const q=h[r.id];r.tx=q?.transaksi||0;if(x.dibagi&&q){r.wj=q.wajib||0;r.w=q.bobot||0}});msg($("#h_msg"),`Memuat perhitungan tersimpan tahun ${x.tahun}${x.dibagi?" (sudah dibagikan "+x.tglBagi+")":""}.`)}}catch(e){}}
$("#h_thn").onchange=async()=>{dasarShu();await muatShu();renderShu()};
function renderShu(){$("#hb").innerHTML=SHUROWS.map(r=>`<tr><td>${esc(r.nama)}</td><td class="n">${rp(r.suk)}</td><td class="n">${idn(r.w)}</td><td class="n"><input type="number" min="0" class="tx" data-id="${esc(r.id)}" value="${r.tx||""}" style="width:130px;text-align:right" aria-label="Transaksi ${esc(r.nama)}"></td>${["hs","ht","ha","hz"].map(p=>`<td class="n" id="${p}_${esc(r.id)}"></td>`).join("")}<td class="n">${rp(r.wj)}</td><td class="n" id="hd_${esc(r.id)}"></td><td><button class="btn ghost sm" data-slip="${esc(r.id)}">Slip</button></td></tr>`).join("");hitungShu()}
$("#hb").oninput=e=>{if(!e.target.classList.contains("tx"))return;const r=SHUROWS.find(x=>x.id===e.target.dataset.id);if(r)r.tx=+e.target.value||0;hitungShu()};
$("#h_laba").oninput=()=>hitungShu();$("#h_pen").oninput=()=>hitungShu();
function hitungShu(){const th=thSHU(),L=+$("#h_laba").value||0,pen=$("#h_pen").value.split("\n").map(s=>s.trim()).filter(Boolean);
 const bagiCad=$("#h_cad").value==="bagi",cc=bagiCad?SHU_PCT.cadangan:0,EP={...SHU_PCT,cadangan:bagiCad?0:SHU_PCT.cadangan,pengurus:SHU_PCT.pengurus+cc*.3,anggota:SHU_PCT.anggota+cc*.2,transaksi:SHU_PCT.transaksi+cc*.2,sukarela:SHU_PCT.sukarela+cc*.3};EPERSEN=EP;const P=Object.fromEntries(Object.entries(EP).map(([k,v])=>[k,Math.round(L*v/100)])),sm=k=>SHUROWS.reduce((t,r)=>t+r[k],0),tW=sm("w"),tT=sm("tx"),n=SHUROWS.length;
 SHUROWS.forEach(r=>{r.hs=tW?P.sukarela*r.w/tW:0;r.ht=tT?P.transaksi*r.tx/tT:0;r.ha=n?P.anggota/n:0;r.z=r.hs+r.ht+r.ha;r.d=r.z+r.wj;
  for(const [p,k] of [["hs","hs"],["ht","ht"],["ha","ha"],["hz","z"],["hd","d"]]){const c=$(`#${p}_${r.id}`);if(c)c.textContent=rp(r[k])}});
 const pw=renderPeng(P.pengurus,pen);
 $("#h_cards").innerHTML=[[bagiCad?"Dana cadangan (dibagikan seluruhnya ke SHU)":"Dana cadangan tetap ("+SHU_PCT.cadangan+"%)",bagiCad?Math.round(L*SHU_PCT.cadangan/100):P.cadangan],["Dana pendidikan & sosial ("+SHU_PCT.pendidikan+"%)",P.pendidikan],["Pengurus & pengawas ("+Math.round(EP.pengurus*10)/10+"%)",P.pengurus],["SHU anggota ("+Math.round(EP.anggota*10)/10+"%)",P.anggota],["SHU transaksi anggota ("+Math.round(EP.transaksi*10)/10+"%)",P.transaksi],["SHU simpanan sukarela ("+Math.round(EP.sukarela*10)/10+"%)",P.sukarela]].map(([l,v])=>`<div class="stat"><b>${rp(v)}</b><span>${l}</span></div>`).join("");
 const st=TBD[th],w=[st?(st.dibagi?`SHU tahun buku ${th} sudah dibagikan pada ${st.tglBagi}.`:`Tahun buku ${th} sudah ditutup. Setelah laba dipastikan, simpan perhitungan lalu bagikan.`):`Tahun buku ${th} belum ditutup. Tutup buku dulu di menu Tutup buku sebelum membagikan.`];
 if(L>0&&!tW)w.push("Belum ada simpanan sukarela pada tahun ini, jadi bagian 10% belum bisa dibagikan. SHU sukarela dihitung dari saldo sukarela setiap akhir bulan (Rp x bulan).");if(L>0&&!tT)w.push("Isi kolom transaksi anggota agar bagian 15% bisa dibagikan.");if(L>0&&!pen.length)w.push("Isi penerima bagian pengurus & pengawas.");if(pw)w.push(pw);
 $("#h_note").innerHTML=w.map((t,i)=>`<div class="msg${i?" err":""}">${esc(t)}</div>`).join("");
 $("#hf").innerHTML=`<tr><td>Jumlah</td><td class="n">${rp(sm("suk"))}</td><td class="n">${idn(tW)}</td><td class="n">${rp(tT)}</td><td class="n">${rp(sm("hs"))}</td><td class="n">${rp(sm("ht"))}</td><td class="n">${rp(sm("ha"))}</td><td class="n">${rp(sm("z"))}</td><td class="n">${rp(sm("wj"))}</td><td class="n">${rp(sm("d"))}</td><td></td></tr>`;
 $("#h_xls").onclick=()=>unduh(SHUROWS.map(r=>({Anggota:r.nama,"Simpanan sukarela":r.suk,"Bobot Rp x bulan":Math.round(r.w),Transaksi:r.tx,"SHU sukarela":Math.round(r.hs),"SHU transaksi":Math.round(r.ht),"SHU anggota":Math.round(r.ha),"Total SHU":Math.round(r.z),"Simpanan wajib":r.wj,"Total diterima":Math.round(r.d)})),`SHU-${th}.xlsx`)}
const rowSave=r=>({id:r.id,nama:r.nama,instansi:r.instansi,sukarela:r.suk,bobot:Math.round(r.w),transaksi:r.tx,wajib:r.wj,shuSukarela:Math.round(r.hs),shuTransaksi:Math.round(r.ht),shuAnggota:Math.round(r.ha),total:Math.round(r.z)});
$("#h_save").onclick=async()=>{const m=$("#h_msg"),L=+$("#h_laba").value,thn=String(thSHU()),pen=$("#h_pen").value.split("\n").map(s=>s.trim()).filter(Boolean);
 if(!(L>0)){msg(m,"Isi laba bersih terlebih dahulu.",true);return}
 try{const lama=await getDoc(doc(db,"shu",thn));if(lama.exists()&&lama.data().dibagi){msg(m,"SHU tahun ini sudah dibagikan, perhitungan tidak bisa diubah.",true);return}
  const data={tahun:+thn,laba:L,persen:EPERSEN,cadangan:$("#h_cad").value,cadanganJumlah:Math.round(L*SHU_PCT.cadangan/100),pengurus:PENG,penerima:pen,hasil:SHUROWS.map(rowSave),disimpan:new Date().toISOString(),oleh:me.nama};
  const b=writeBatch(db);b.set(doc(db,"shu",thn),data);logOp(b,{aksi:lama.exists()?"ubah":"tambah",kol:"shu",id:thn,sebelum:lama.exists()?{laba:lama.data().laba}:null,sesudah:{laba:L},detail:`SHU tahun ${thn}`,alasan:lama.exists()?"Perhitungan ulang":""});await b.commit();
  msg(m,`Perhitungan SHU tahun ${thn} tersimpan.`)}catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};
const slipHtml=(r,th)=>`<section style="page-break-after:always;font:14px/1.5 system-ui,sans-serif;max-width:560px;margin:24px auto;border:2px solid #104020;padding:22px"><div style="text-align:center"><b style="font-size:17px">KOPERASI PGRI SANGGAM BARINTAK</b><br>Jalan Ramania, Gedung SKB, Tanjung Redeb, Berau<hr><b>SLIP PEMBAGIAN SISA HASIL USAHA</b><br><span style="background:#104020;color:#fff;padding:2px 12px;border-radius:12px;display:inline-block;margin-top:6px">TAHUN BUKU ${th} (1 Januari - 31 Desember ${th})</span></div><p>Nama: <b>${esc(r.nama)}</b><br>Instansi: ${esc(r.instansi||"-")}</p><table style="width:100%;border-collapse:collapse">${[["Simpanan wajib tahun "+th+" (dikembalikan)",r.wj],["SHU simpanan sukarela",r.hs],["SHU transaksi",r.ht],["SHU anggota",r.ha]].map(([l,v])=>`<tr><td style="border-bottom:1px solid #ccc;padding:5px 0">${l}</td><td style="text-align:right;border-bottom:1px solid #ccc">${rp(v)}</td></tr>`).join("")}<tr><td style="padding:8px 0"><b>TOTAL DITERIMA</b></td><td style="text-align:right"><b>${rp(r.wj+r.z)}</b></td></tr></table><p style="font-size:12px;color:#555">SHU sukarela dihitung dari lama simpanan: ${idn(r.w)} (Rp x bulan). Simpanan pokok dan sukarela tidak ikut dikembalikan; simpanan sukarela dapat ditarik kapan saja.</p><p style="text-align:right">Dicetak ${today()}<br><br><br>Bendahara</p></section>`;
const slipPeng=(o,th)=>`<section style="page-break-after:always;font:14px/1.5 system-ui,sans-serif;max-width:560px;margin:24px auto;border:2px solid #104020;padding:22px"><div style="text-align:center"><b style="font-size:17px">KOPERASI PGRI SANGGAM BARINTAK</b><br>Jalan Ramania, Gedung SKB, Tanjung Redeb, Berau<hr><b>SLIP SHU PENGURUS & PENGAWAS</b><br><span style="background:#104020;color:#fff;padding:2px 12px;border-radius:12px;display:inline-block;margin-top:6px">TAHUN BUKU ${th} (1 Januari - 31 Desember ${th})</span></div><p>Nama / jabatan: <b>${esc(o.nama)}</b></p><table style="width:100%;border-collapse:collapse"><tr><td style="border-bottom:1px solid #ccc;padding:5px 0">Bagian dari SHU pengurus & pengawas</td><td style="text-align:right;border-bottom:1px solid #ccc">${o.persen}%</td></tr><tr><td style="padding:8px 0"><b>TOTAL DITERIMA</b></td><td style="text-align:right"><b>${rp(o.jumlah)}</b></td></tr></table><p style="text-align:right">Dicetak ${today()}<br><br><br>Bendahara</p></section>`;
const slipDoc=(rows,th,peng=[])=>`<!doctype html><meta charset="utf-8"><title>Slip SHU tahun buku ${th}</title><style>@media print{.np{display:none}}</style><div class="np" style="text-align:center;margin:12px"><button onclick="print()">Cetak / simpan PDF</button></div>`+rows.map(r=>slipHtml(r,th)).join("")+peng.map(o=>slipPeng(o,th)).join("");
$("#hb").onclick=e=>{const b=e.target.closest("[data-slip]");if(!b)return;const r=SHUROWS.find(x=>x.id===b.dataset.slip);if(r)window.open(URL.createObjectURL(new Blob([slipDoc([r],thSHU())],{type:"text/html"})),"_blank")};
$("#h_slip").onclick=()=>{const th=thSHU(),a=document.createElement("a");a.href=URL.createObjectURL(new Blob([slipDoc(SHUROWS,th,PENG)],{type:"text/html"}));a.download=`slip-SHU-tahun-buku-${th}.html`;a.click()};
async function syncSemua(){await getAnggota(true);await muatTB().catch(()=>{});const W={};const s=await getDocs(collection(db,"simpanan")),t={};s.forEach(d=>{const x=d.data();(W[x.anggotaId]??=[]).push(x);(t[x.anggotaId]??={pokok:0,wajib:0,sukarela:0})[x.jenis]+=x.jumlah});
 const aktif=ANG.filter(a=>a.status!=="berhenti"),g={pokok:0,wajib:0,sukarela:0};
 for(let i=0;i<aktif.length;i+=400){const b=writeBatch(db);aktif.slice(i,i+400).forEach(a=>{const x=t[a.id]||{pokok:0,wajib:0,sukarela:0};for(const k in g)g[k]+=x[k];b.set(doc(db,"publik_anggota",a.id),{no:a.no,nama:a.nama,instansi:a.instansi||"",...x,total:x.pokok+x.wajib+x.sukarela,...pubW(a,W[a.id]||[])},{merge:true})});await b.commit()}
 await setDoc(doc(db,"publik_ringkasan","utama"),{jumlahAnggota:aktif.length,...g,tahunDibagi:Object.keys(TBD).filter(y=>TBD[y].dibagi).map(Number),diperbarui:new Date().toISOString()},{merge:true})}
$("#h_bagi").onclick=async()=>{const th=thSHU(),m=$("#h_msg");
 try{const lama=await getDoc(doc(db,"shu",String(th)));await muatTB();
  if(!TBD[th])return msg(m,`Tutup tahun buku ${th} dulu di menu Tutup buku.`,true);if(!lama.exists())return msg(m,"Simpan perhitungan SHU dulu.",true);if(lama.data().dibagi)return msg(m,"SHU tahun ini sudah dibagikan.",true);
  const v=await tanya({judul:`Bagikan SHU tahun buku ${th}`,tombol:"Bagikan",fields:[{type:"note",label:`Simpanan wajib tahun ${th} seluruh anggota dikembalikan dan dihapus dari rekap simpanan, lalu SHU dibayarkan. Total kas keluar ${rp(SHUROWS.reduce((t,r)=>t+r.wj+Math.round(r.z),0)+PENG.reduce((t,o)=>t+o.jumlah,0))}. Tanggal pembagian harus di tahun buku yang masih terbuka.`},{k:"tgl",label:"Tanggal pembagian",type:"date",req:1}],nilai:{tgl:today()},validasi:o=>o.tgl<`${th+1}-01-01`?`Tanggal pembagian harus setelah 31 Desember ${th}.`:tertutup(o.tgl)?"Tahun buku pada tanggal itu sudah ditutup.":""});if(!v)return;
  const ops=[];let urut=Date.now(),wt=0;
  SHUROWS.forEach(r=>{const z=Math.round(r.z);if(r.wj>0){wt+=r.wj;ops.push(b=>b.set(doc(db,"bku",`shu${th}_${r.id}_w`),{tanggal:v.tgl,kode:"SIM",keterangan:`Pengembalian simpanan wajib ${th} an: ${r.nama}`,kredit:0,debit:r.wj,ref:`shu${th}_${r.id}`,anggotaId:r.id,jenis:"wajib",urut:urut++}))}
   if(z>0)ops.push(b=>b.set(doc(db,"bku",`shu${th}_${r.id}_s`),{tanggal:v.tgl,kode:"SHU",keterangan:`Pembagian SHU tahun buku ${th} an: ${r.nama}`,kredit:0,debit:z,urut:urut++}))});
  PENG.forEach((o,i)=>{if(o.jumlah>0)ops.push(b=>b.set(doc(db,"bku",`shu${th}_peng${i}`),{tanggal:v.tgl,kode:"SHU",keterangan:`SHU pengurus/pengawas tahun buku ${th} an: ${o.nama} (${o.persen}%)`,kredit:0,debit:o.jumlah,urut:urut++}))});
  const Lb=+$("#h_laba").value,cj=Math.round(Lb*SHU_PCT.cadangan/100),modeCad=$("#h_cad").value;
  if(modeCad!=="bagi"&&cj>0)ops.push(b=>b.set(doc(db,"bku",`shu${th}_cadangan`),{tanggal:v.tgl,kode:"CAD",keterangan:`Dana cadangan tahun buku ${th} sebesar ${rp(cj)} tetap menjadi dana cadangan koperasi (pencatatan terpisah, tidak mengubah kas)`,kredit:0,debit:0,nominal:cj,urut:urut++}));
  SIM.filter(x=>x.jenis==="wajib"&&(x.periode||"").startsWith(th+"-")).forEach(x=>ops.push(b=>b.delete(doc(db,"simpanan",x.id))));
  for(let i=0;i<ops.length;i+=400){const b=writeBatch(db);ops.slice(i,i+400).forEach(f=>f(b));await b.commit()}
  const b=writeBatch(db);b.set(doc(db,"shu",String(th)),{dibagi:true,tglBagi:v.tgl,wajibKembali:wt,laba:Lb,persen:EPERSEN,cadangan:modeCad,cadanganJumlah:cj,pengurus:PENG,hasil:SHUROWS.map(rowSave)},{merge:true});b.set(doc(db,"tahunbuku",String(th)),{dibagi:true,tglBagi:v.tgl,wajibKembali:wt},{merge:true});
  logOp(b,{aksi:"bagi_shu",kol:"shu",id:String(th),sesudah:{tglBagi:v.tgl,wajibKembali:wt},detail:`Pembagian SHU tahun buku ${th}`});await b.commit();
  await syncSemua();await muatTB();await loadShu();msg(m,`SHU tahun buku ${th} dibagikan. Simpanan wajib tahun itu sudah dihapus dari rekap anggota dan pengeluaran tercatat di Buku kas.`)}
 catch(e){msg(m,"Gagal: "+e.message,true)}};

// ---------- Impor Excel ----------
$("#imp").onclick=async()=>{const m=$("#imsg"),file=$("#f").files[0];if(!file){msg(m,"Pilih file Excel dulu.",true);return}
 $("#imp").disabled=true;msg(m,"Membaca file…");
 try{const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true}),wsS=wb.Sheets["SIMPANAN ANGGOTA"],wsB=wb.Sheets["BKU"];
  if(!wsS&&!wsB)throw new Error('Sheet "SIMPANAN ANGGOTA" atau "BKU" tidak ditemukan');
  const ops=[],TAHUN=2026;let nA=0,nB=0;
  if(wsS){const rows=XLSX.utils.sheet_to_json(wsS,{header:1,defval:null}).filter(r=>typeof r[0]==="number"&&typeof r[2]==="string"&&r[2].trim());const tot={pokok:0,wajib:0,sukarela:0};
   for(const r of rows){const no=r[0],id="A"+String(no).padStart(3,"0"),tgl=r[1] instanceof Date?ymd(r[1]):"",pokok=+r[4]||0,suk=+r[5]||0;let wajib=0;const wb=[];
    ops.push([doc(db,"anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),tglDaftar:tgl}]);
    if(pokok)ops.push([doc(db,"simpanan",id+"_pokok"),{anggotaId:id,jenis:"pokok",jumlah:pokok,tanggal:tgl}]);
    if(suk)ops.push([doc(db,"simpanan",id+"_sukarela"),{anggotaId:id,jenis:"sukarela",jumlah:suk,tanggal:tgl}]);
    for(let b=0;b<12;b++){const v=+r[6+b]||0;if(v){wajib+=v;wb.push(`${TAHUN}-${String(b+1).padStart(2,"0")}`);ops.push([doc(db,"simpanan",`${id}_wajib_${TAHUN}_${String(b+1).padStart(2,"0")}`),{anggotaId:id,jenis:"wajib",jumlah:v,periode:`${TAHUN}-${String(b+1).padStart(2,"0")}`}])}}
    tot.pokok+=pokok;tot.wajib+=wajib;tot.sukarela+=suk;
    ops.push([doc(db,"publik_anggota",id),{no,nama:r[2].trim(),instansi:(r[3]||"").trim(),pokok,wajib,sukarela:suk,total:pokok+wajib+suk,wajibBulan:wb,daftar:tgl.slice(0,7)}])}
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

// ---------- Laporan keuangan ----------
const AKUN=[["","Otomatis menurut kode"],["pend_usaha","Pendapatan usaha"],["pend_lain","Pendapatan lain (jasa bank, dll)"],["beban_usaha","Beban/pembelian unit usaha"],["beban_ops","Beban operasional & administrasi"],["aset_tetap","Pembelian aset tetap/inventaris"],["piutang","Piutang/pinjaman (keluar = disalurkan, masuk = dilunasi)"],["utang","Utang (masuk = diterima, keluar = dibayar)"],["dist_shu","SHU/cadangan yang dibayarkan"],["simpanan","Simpanan anggota"]];
const ALBL=Object.fromEntries(AKUN),PEND=["pend_usaha","pend_lain"],BEB=["beban_usaha","beban_ops"];
$("#b_akun").innerHTML=AKUN.map(([v,l])=>`<option value="${v}">${esc(l)}</option>`).join("");
const kurang=x=>{const k=[];if(!x.keterangan)k.push("keterangan");if(!x.kode)k.push("kode");if(x.kode==="SIM"&&!x.ref&&!x.konfirmasi)k.push("anggota & pembagian simpanan");return k};
const net=x=>(x.kredit||0)-(x.debit||0),UNIT=x=>TIPE[x.kode]==="usaha"?x.kode:"UMUM";
const akunDari=x=>{if(x.akun)return x.akun;const t=TIPE[x.kode]||"umum";return t==="usaha"?(x.kredit?"pend_usaha":"beban_usaha"):t==="umum"?(x.kredit?"pend_lain":"beban_ops"):t};
const S=(L,f)=>L.filter(x=>f(akunDari(x))).reduce((t,x)=>t+net(x),0),sm=(L,k)=>L.reduce((t,x)=>t+(x[k]||0),0);
function hitungKeu(th){const s=th+"-01-01",e=th+"-12-31",all=BKU.filter(x=>(x.tanggal||"")<=e),yr=all.filter(x=>x.tanggal>=s),pr=all.filter(x=>x.tanggal<s),is=a=>b=>b===a,P=a=>S(yr,is(a));
 const lb=L=>S(L,a=>PEND.includes(a))+S(L,a=>BEB.includes(a)),k={th,pu:P("pend_usaha"),pl:P("pend_lain"),bu:-P("beban_usaha"),bo:-P("beban_ops")};
 k.pend=k.pu+k.pl;k.beb=k.bu+k.bo;k.laba=k.pend-k.beb;k.labaLalu=lb(pr);
 k.unit=Object.fromEntries([...new Set([...KODE.filter(c=>TIPE[c]==="usaha"),"UMUM"])].map(u=>{const L=yr.filter(x=>UNIT(x)===u);return [u,S(L,a=>PEND.includes(a))+S(L,a=>BEB.includes(a))]}));
 k.kas=all.reduce((t,x)=>t+net(x),0);k.kasAwal=pr.reduce((t,x)=>t+net(x),0);k.aset=-S(all,is("aset_tetap"));k.piutang=-S(all,is("piutang"));k.utang=S(all,is("utang"));k.dist=-S(all,is("dist_shu"));
 const sb=S(all,is("simpanan")),ss={pokok:0,wajib:0,sukarela:0};SIM.forEach(x=>{const d=x.tanggal||(x.periode?x.periode+"-31":"");if(!d||d<=e)ss[x.jenis]+=x.jumlah});
 Object.values(TBD).forEach(t=>{if(t.dibagi&&t.tglBagi>e)ss.wajib+=t.wajibKembali||0});k.ss=ss;k.selisih=sb-ss.pokok-ss.wajib-ss.sukarela;k.ditahan=k.labaLalu-k.dist;k.ekuitas=sb+k.ditahan+k.laba;k.pasiva=k.utang+k.ekuitas;k.aktiva=k.kas+k.piutang+k.aset;
 k.masuk=AKUN.filter(a=>a[0]).map(([a,l])=>[l,yr.filter(x=>akunDari(x)===a).reduce((t,x)=>t+(x.kredit||0),0)]).filter(r=>r[1]);
 k.keluar=AKUN.filter(a=>a[0]).map(([a,l])=>[l,yr.filter(x=>akunDari(x)===a).reduce((t,x)=>t+(x.debit||0),0)]).filter(r=>r[1]);
 k.auto=yr.filter(x=>!x.akun&&x.kode!=="SIM").length;k.simBelum=BKU.filter(x=>x.kode==="SIM"&&!x.ref&&!x.konfirmasi).length;k.tdkLengkap=yr.filter(x=>kurang(x).length).length;return k}
function sect(XL,j,rows){XL.push({Laporan:j,Uraian:"",Jumlah:""});rows.forEach(([l,v,b,i])=>XL.push({Laporan:"",Uraian:(i?"   ":"")+l,Jumlah:v??""}));
 return `<div class="panel scroll"><h2 style="margin-bottom:8px">${esc(j)}</h2><table><tbody>${rows.map(([l,v,b,i])=>`<tr style="${b?"font-weight:700":""}"><td style="${i?"padding-left:28px":""}">${esc(l)}</td><td class="n">${typeof v==="number"?rp(v):esc(v??"")}</td></tr>`).join("")}</tbody></table></div>`}
const LR=k=>[["PENDAPATAN",null,1],["Pendapatan usaha (cafe, konveksi, dll)",k.pu,0,1],["Pendapatan lain",k.pl,0,1],["Jumlah pendapatan",k.pend,1],["BEBAN",null,1],["Beban/pembelian unit usaha",k.bu,0,1],["Beban operasional & administrasi",k.bo,0,1],["Jumlah beban",k.beb,1],["LABA (RUGI) BERSIH TAHUN "+k.th,k.laba,1],["Laba (rugi) per unit",null,1],...Object.entries(k.unit).map(([u,v])=>[u==="UMUM"?"Umum / di luar unit":"Unit "+u,v,0,1])];
const NR=k=>[["ASET",null,1],["Kas dan bank",k.kas,0,1],["Piutang / pinjaman",k.piutang,0,1],["Aset tetap / inventaris (harga perolehan)",k.aset,0,1],["JUMLAH ASET",k.aktiva,1],["KEWAJIBAN",null,1],["Utang",k.utang,0,1],["EKUITAS",null,1],["Simpanan pokok",k.ss.pokok,0,1],["Simpanan wajib",k.ss.wajib,0,1],["Simpanan sukarela",k.ss.sukarela,0,1],["Simpanan belum dirinci (selisih BKU dan data anggota)",k.selisih,0,1],["Cadangan & laba ditahan tahun lalu (setelah SHU dibayarkan)",k.ditahan,0,1],["Laba tahun berjalan",k.laba,0,1],["Jumlah ekuitas",k.ekuitas,1],["JUMLAH KEWAJIBAN + EKUITAS",k.pasiva,1]];
const KS=k=>[["Saldo kas awal tahun",k.kasAwal,1],["Penerimaan",null,1],...k.masuk.map(r=>[r[0],r[1],0,1]),["Jumlah penerimaan",k.masuk.reduce((t,r)=>t+r[1],0),1],["Pengeluaran",null,1],...k.keluar.map(r=>[r[0],r[1],0,1]),["Jumlah pengeluaran",k.keluar.reduce((t,r)=>t+r[1],0),1],["SALDO KAS AKHIR TAHUN",k.kas,1]];
async function muatKeu(){const [b,s]=await Promise.all([getDocs(collection(db,"bku")),getDocs(collection(db,"simpanan")),getAnggota(true)]);
 BKU=b.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(a.tanggal||"").localeCompare(b.tanggal||"")||(a.urut||0)-(b.urut||0));SIM=s.docs.map(d=>({id:d.id,...d.data()}))}
async function loadLaporan(){if(!$("#k_thn").value)$("#k_thn").value=new Date().getFullYear();await muatKeu();await renderLap()}
async function renderLap(){const th=+$("#k_thn").value,j=$("#k_jenis").value,k=hitungKeu(th),XL=[],o=[];
 const w=[];if(k.simBelum)w.push(`${k.simBelum} transaksi SIM di Buku kas belum dikonfirmasi.`);if(k.tdkLengkap)w.push(`${k.tdkLengkap} transaksi tahun ${th} belum lengkap isiannya.`);if(k.auto)w.push(`${k.auto} transaksi memakai kategori otomatis menurut kode; periksa lewat Buku kas > Ubah bila perlu.`);if(Math.abs(k.selisih)>=1)w.push(`Total simpanan di BKU berbeda ${rp(k.selisih)} dari data simpanan anggota.`);
 if(Math.abs(k.aktiva-k.pasiva)>=1)w.push("Neraca belum seimbang, periksa kategori transaksi.");
 o.push(w.length?`<div class="warn">${w.map(t=>`<span>⚠ ${esc(t)}</span>`).join("")}</div>`:`<div class="warn"><span>Data lengkap dan neraca seimbang.</span></div>`);
 if(j==="tahunan"){const act=ANG.filter(a=>a.status!=="berhenti").length,baru=ANG.filter(a=>(a.tglDaftar||"").startsWith(th)).length;let shu=null;try{const d=await getDoc(doc(db,"shu",String(th)));if(d.exists())shu=d.data()}catch(e){}
  o.push(`<div class="panel" style="text-align:center"><h2>LAPORAN TAHUNAN ${th}</h2><p style="margin:4px 0 0">Koperasi PGRI Sanggam Barintak<br>Jalan Ramania, Gedung SKB, Tanjung Redeb, Berau</p></div>`);
  o.push(sect(XL,"I. Keanggotaan & simpanan",[["Anggota aktif",String(act),0,1],["Anggota baru tahun "+th,String(baru),0,1],["Simpanan pokok",k.ss.pokok,0,1],["Simpanan wajib",k.ss.wajib,0,1],["Simpanan sukarela",k.ss.sukarela,0,1],["Total simpanan anggota",k.ss.pokok+k.ss.wajib+k.ss.sukarela,1]]));
  o.push(sect(XL,"II. Laporan laba rugi",LR(k)),sect(XL,"III. Neraca per 31 Desember "+th,NR(k)),sect(XL,"IV. Arus kas",KS(k)));
  o.push(sect(XL,"V. Sisa hasil usaha (SHU)",shu?[["Laba bersih",shu.laba,1],...Object.entries(shu.persen||{}).map(([n,p])=>[`${n} (${Math.round(p*10)/10}%)`,Math.round(shu.laba*p/100),0,1])]:[["Perhitungan SHU tahun ini belum disimpan di menu SHU.",null]]));
  o.push(`<div class="panel" style="display:flex;justify-content:space-around;text-align:center;flex-wrap:wrap;gap:30px"><div>Ketua<br><br><br><br><b>Marsono</b></div><div>Bendahara<br><br><br><br><b>Eli Sutrisnaniati</b></div><div>Ketua Pengawas<br><br><br><br><b>Hj. Mardiatul Idalisah, S.Pd., M.M.</b></div></div>`)}
 else o.push(sect(XL,{lr:"Laporan laba rugi tahun "+th,neraca:"Neraca per 31 Desember "+th,kas:"Laporan arus kas tahun "+th}[j],{lr:LR,neraca:NR,kas:KS}[j](k)));
 $("#k_out").innerHTML=o.join("");$("#k_xls").onclick=()=>unduh(XL,`laporan-${j}-${th}.xlsx`)}
$("#k_thn").onchange=renderLap;$("#k_jenis").onchange=renderLap;
$("#h_ambil").onclick=async()=>{if(!can("bku","lihat"))return alert("Butuh hak lihat Buku kas.");await muatKeu();const k=hitungKeu(+$("#h_thn").value);$("#h_laba").value=Math.max(0,Math.round(k.laba));hitungShu();msg($("#h_msg"),`Laba bersih tahun ${k.th} dari laporan laba rugi: ${rp(k.laba)}.`)};

// ---------- Tahun buku, kode transaksi, rekap RAT ----------
let TB=new Set(),TBD={};
async function muatTB(){const s=await getDocs(collection(db,"tahunbuku"));TBD={};s.docs.forEach(d=>TBD[d.id]=d.data());TB=new Set(Object.keys(TBD))}
const tertutup=t=>!!t&&TB.has(String(t).slice(0,4));
function renderKode(){$("#b_kode").innerHTML=KODE.map(k=>`<option value="${esc(k)}">${esc(k)} (${esc(KLBL[k]||"")})</option>`).join("")}
async function muatKode(){renderKode();const s=await getDocs(collection(db,"kode_bku"));s.docs.forEach(d=>{const x=d.data();if(!KODE.includes(x.kode))KODE.push(x.kode);TIPE[x.kode]=x.tipe;KLBL[x.kode]=x.label});renderKode()}
$("#b_kodenew").onclick=async()=>{try{const v=await tanya({judul:"Tambah kode transaksi",alasan:false,fields:[{k:"kode",label:"Kode (huruf/angka tanpa spasi, mis. SEWA)",req:1},{k:"label",label:"Nama kode",req:1},{k:"tipe",label:"Perlakuan dalam laporan",type:"select",opt:[["usaha","Unit usaha (masuk = pendapatan, keluar = beban; laba dihitung per unit)"],["umum","Umum (masuk = pendapatan lain, keluar = beban operasional)"],["piutang","Piutang / peminjaman dana (keluar = dipinjamkan, masuk = dikembalikan)"],["utang","Utang"],["aset_tetap","Pembelian aset tetap"]]}],validasi:o=>/^[A-Z0-9]{2,12}$/.test(o.kode.toUpperCase())?(KODE.includes(o.kode.toUpperCase())?"Kode sudah ada.":""):"Kode 2-12 huruf/angka tanpa spasi."});if(!v)return;
 const k=v.kode.toUpperCase(),d={kode:k,label:v.label,tipe:v.tipe},b=writeBatch(db);b.set(doc(db,"kode_bku",k),d);logOp(b,{aksi:"tambah",kol:"kode_bku",id:k,sesudah:d,detail:k});await b.commit();await muatKode();$("#b_kode").value=k}catch(e){alert("Gagal: "+e.message)}};
function piutangPihak(th){const m={};BKU.filter(x=>(x.tanggal||"")<=th+"-12-31"&&akunDari(x)==="piutang").forEach(x=>{const p=x.pihak||"(tanpa nama)";m[p]=(m[p]||0)+(x.debit||0)-(x.kredit||0)});return Object.entries(m).filter(r=>Math.abs(r[1])>=1)}
async function loadTutup(){if(!$("#c_thn").value)$("#c_thn").value=new Date().getFullYear();await muatKeu();await muatTB();renderTutup()}
function renderTutup(){const th=+$("#c_thn").value,k=hitungKeu(th),t=TBD[th],pp=piutangPihak(th),XL=[],pc=SHU_PCT.anggota+SHU_PCT.transaksi+SHU_PCT.sukarela,
 wj=SIM.filter(x=>x.jenis==="wajib"&&(x.periode||"").startsWith(th+"-")).reduce((s,x)=>s+x.jumlah,0),shuA=Math.max(0,Math.round(k.laba*pc/100)),butuh=wj+shuA,sisa=k.kas-butuh;
 $("#c_tutup").classList.toggle("hide",!!t||!can("tutup","ubah"));$("#c_buka").classList.toggle("hide",!t||!!t.dibagi||!can("tutup","ubah"));
 $("#c_out").innerHTML=`<div class="warn"><span><b>Tahun buku ${th}: ${t?(t.dibagi?"DITUTUP, SHU sudah dibagikan":"DITUTUP"):"TERBUKA"}</b></span></div>`
 +sect(XL,"Pemeriksaan sebelum tutup buku",[["Transaksi SIM belum dikonfirmasi",String(k.simBelum),0,1],["Transaksi tahun ini belum lengkap",String(k.tdkLengkap),0,1],["Neraca seimbang",Math.abs(k.aktiva-k.pasiva)<1?"Ya":"Tidak",0,1],["Selisih simpanan BKU vs data anggota",k.selisih,0,1]])
 +sect(XL,"Sisa piutang (peminjaman dana) per 31 Desember "+th,[...pp.map(([n,v])=>[n,v,0,1]),["Total sisa piutang",k.piutang,1]])
 +sect(XL,"Perkiraan SHU yang dapat dibagikan",[["Laba bersih tahun "+th,k.laba,1],[`SHU untuk anggota (${pc}% laba)`,shuA,0,1],["Simpanan wajib tahun ini yang dikembalikan",wj,0,1],["Kebutuhan kas pembagian",butuh,1],["Saldo kas per 31 Desember",k.kas,0,1],["Sisa kas setelah pembagian",sisa,1],[sisa<0?"Kas belum cukup: tunggu pelunasan piutang atau sesuaikan pembagian.":"Kas mencukupi untuk pembagian.",null]])}
$("#c_thn").onchange=renderTutup;
$("#c_tutup").onclick=async()=>{const th=+$("#c_thn").value,m=$("#c_msg");if(today()<th+"-12-01")return msg(m,`Tahun buku baru bisa ditutup mulai 1 Desember ${th}.`,true);
 const k=hitungKeu(th),v=await tanya({judul:"Tutup tahun buku "+th,tombol:"Tutup buku",fields:[{type:"note",label:`Setelah ditutup, transaksi bertanggal ${th} tidak bisa ditambah, diubah, atau dihapus. Laba bersih ${rp(k.laba)}, sisa piutang ${rp(k.piutang)}.`}]});if(!v)return;
 const d={tahun:th,status:"tutup",laba:Math.round(k.laba),piutang:Math.round(k.piutang),kas:Math.round(k.kas),ditutupOleh:me.nama,waktu:new Date().toISOString()},b=writeBatch(db);b.set(doc(db,"tahunbuku",String(th)),d);logOp(b,{aksi:"tutup_buku",kol:"tahunbuku",id:String(th),sesudah:d,alasan:v.alasan,detail:"Tahun buku "+th});
 try{await b.commit();await muatTB();renderTutup();msg(m,`Tahun buku ${th} ditutup.`)}catch(e){msg(m,"Gagal: "+e.message,true)}};
$("#c_buka").onclick=async()=>{const th=+$("#c_thn").value,m=$("#c_msg"),v=await tanya({judul:"Buka kembali tahun buku "+th,tombol:"Buka kembali",fields:[{type:"note",label:"Transaksi tahun ini bisa diubah lagi. Hanya bisa dilakukan sebelum SHU dibagikan."}]});if(!v)return;
 const b=writeBatch(db);b.delete(doc(db,"tahunbuku",String(th)));logOp(b,{aksi:"buka_buku",kol:"tahunbuku",id:String(th),alasan:v.alasan,detail:"Buka kembali tahun buku "+th});
 try{await b.commit();await muatTB();renderTutup();msg(m,`Tahun buku ${th} dibuka kembali.`)}catch(e){msg(m,"Gagal: "+e.message,true)}};
$("#k_rat").onclick=async()=>{const th=+$("#k_thn").value;await muatKeu();await muatTB();const k=hitungKeu(th),wb=XLSX.utils.book_new(),add=(n,r)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(r),n),X=r=>{const L=[];sect(L,"",r);return L.slice(1)};
 const sd=await getDoc(doc(db,"shu",String(th))),act=ANG.filter(a=>a.status!=="berhenti"),t={};SIM.forEach(x=>{(t[x.anggotaId]??={pokok:0,wajib:0,sukarela:0})[x.jenis]+=x.jumlah});
 add("Ringkasan",X([["Tahun buku",String(th)],["Anggota aktif",String(act.length)],["Anggota baru",String(ANG.filter(a=>(a.tglDaftar||"").startsWith(th)).length)],["Total simpanan pokok",k.ss.pokok],["Total simpanan wajib",k.ss.wajib],["Total simpanan sukarela",k.ss.sukarela],["Laba bersih",k.laba],["Saldo kas",k.kas],["Sisa piutang",k.piutang],["Status tahun buku",TBD[th]?(TBD[th].dibagi?"Ditutup, SHU dibagikan":"Ditutup"):"Terbuka"]]));
 add("Laba Rugi",X(LR(k)));add("Neraca",X(NR(k)));add("Arus Kas",X(KS(k)));add("Piutang",piutangPihak(th).map(([n,v])=>({Peminjam:n,"Sisa piutang":v})));
 add("SHU per anggota",(sd.exists()?sd.data().hasil||[]:[]).map(r=>({Anggota:r.nama,"Bobot sukarela (Rp x bulan)":r.bobot||0,"SHU sukarela":r.shuSukarela,"SHU transaksi":r.shuTransaksi,"SHU anggota":r.shuAnggota,"Simpanan wajib":r.wajib||0,"Total diterima":(r.wajib||0)+r.total})));
 add("Simpanan anggota",act.map(a=>{const x=t[a.id]||{pokok:0,wajib:0,sukarela:0};return {No:a.no,Nama:a.nama,Instansi:a.instansi||"",Pokok:x.pokok,Wajib:x.wajib,Sukarela:x.sukarela,Total:x.pokok+x.wajib+x.sukarela}}));
 XLSX.writeFile(wb,`rekap-RAT-${th}.xlsx`)};

// ---------- Pengurus/pengawas & dana cadangan ----------
let PENG=[],EPERSEN=SHU_PCT;ALBL.memo="Catatan dana cadangan (tidak mengubah kas)";
const persenPeng=l=>{const m=l.match(/\|\s*(\d+(?:[.,]\d+)?)\s*$/);if(m)return +m[1].replace(",",".");return /pengawas/i.test(l)?(/ketua/i.test(l)?10:5):20},namaPeng=l=>l.replace(/\|\s*[\d.,]+\s*$/,"").trim();
$("#h_cards").insertAdjacentHTML("beforebegin",`<div class="fg"><div><label for="h_cad">Dana cadangan</label><select id="h_cad"><option value="tetap">Tetap menjadi dana cadangan (dicatat terpisah di BKU)</option><option value="bagi">Dibagikan seluruhnya sebagai SHU (pengurus & pengawas +30%, anggota +20%, transaksi +20%, sukarela +30%)</option></select></div></div><small style="color:var(--muted)">Baris pengurus/pengawas: Nama (Jabatan). Ketua, Sekretaris, Bendahara, dan Koordinator Unit Usaha masing-masing 20%; Ketua Pengawas 10%; Anggota Pengawas masing-masing 5%. Untuk persen lain tulis: Nama (Jabatan) | 15</small>`);
$("#h_cards").insertAdjacentHTML("afterend",'<div id="h_peng" class="scroll" style="margin-top:12px"></div>');
$("#h_cad").onchange=()=>hitungShu();
function renderPeng(tot,pen){PENG=pen.map(l=>{const p=persenPeng(l);return {nama:namaPeng(l),persen:p,jumlah:Math.round(tot*p/100)}});const sp=PENG.reduce((t,o)=>t+o.persen,0);
 $("#h_peng").innerHTML=`<h3 style="margin:8px 0">Pembagian pengurus & pengawas (${rp(tot)})</h3><table><thead><tr><th>Nama / jabatan</th><th class="n">Persen</th><th class="n">Jumlah</th><th>Slip</th></tr></thead><tbody>${PENG.map((o,i)=>`<tr><td>${esc(o.nama)}</td><td class="n">${o.persen}%</td><td class="n">${rp(o.jumlah)}</td><td><button class="btn ghost sm" data-pslip="${i}">Slip</button></td></tr>`).join("")}<tr style="font-weight:700"><td>Jumlah</td><td class="n">${sp}%</td><td class="n">${rp(PENG.reduce((t,o)=>t+o.jumlah,0))}</td><td></td></tr></tbody></table>`;
 return pen.length&&Math.abs(sp-100)>0.01?`Total persen pengurus & pengawas ${sp}%, seharusnya 100%.`:""}
$("#h_peng").onclick=e=>{const b=e.target.closest("[data-pslip]");if(!b)return;const o=PENG[+b.dataset.pslip];if(o)window.open(URL.createObjectURL(new Blob([slipDoc([],thSHU(),[o])],{type:"text/html"})),"_blank")};

// ---------- Sinkron data portal anggota ----------
$("#t_cards").insertAdjacentHTML("beforebegin",'<div class="row noprint"><button class="btn ghost" id="t_pub">Perbarui data portal anggota (rekap wajib bulanan & tunggakan)</button></div><div id="t_pubmsg" class="msg hide"></div>');
$("#t_pub").classList.toggle("hide",false);
$("#t_pub").onclick=async()=>{const m=$("#t_pubmsg");if(!(can("simpanan","ubah")||can("anggota","ubah")))return msg(m,"Butuh hak ubah Simpanan atau Anggota.",true);msg(m,"Memperbarui…");try{await syncSemua();msg(m,"Data portal anggota diperbarui.")}catch(e){msg(m,"Gagal: "+e.message,true)}};
