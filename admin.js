import {initializeApp,deleteApp} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {getAuth,signInWithEmailAndPassword,signOut,onAuthStateChanged,createUserWithEmailAndPassword,sendPasswordResetEmail} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {getFirestore,collection,getDocs,doc,getDoc,setDoc,addDoc,writeBatch,serverTimestamp,query,where,orderBy,limit} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import {firebaseConfig} from "./firebase-config.js";
const app=initializeApp(firebaseConfig),auth=getAuth(app),db=getFirestore(app);
const $=s=>document.querySelector(s),rp=n=>"Rp"+Math.round(n||0).toLocaleString("id-ID");
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const ymd=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const today=()=>ymd(new Date());
const msg=(el,t,err)=>{el.textContent=t;el.classList.remove("hide");el.classList.toggle("err",!!err)};
const unduh=(rows,nama)=>{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Data");XLSX.writeFile(wb,nama)};
const MOD=[["anggota","Anggota"],["simpanan","Simpanan"],["bku","Buku kas (BKU)"],["surat","Surat masuk/keluar"],["shu","SHU"],["log","Log aktivitas"]];
const LBL=Object.fromEntries(MOD);
let me=null,ANG=[],SIM=[],BKU=[],SURAT=[],SHUROWS=[],USERS=[],editing=null;
const can=(m,k)=>!!me&&(me.master||(me[k]||[]).includes(m));
async function catat(aksi,detail){try{await addDoc(collection(db,"log"),{uid:auth.currentUser.uid,nama:me.nama,aksi,detail,waktu:serverTimestamp()})}catch(e){console.warn("log gagal",e)}}
async function getAnggota(force){if(force||!ANG.length){const s=await getDocs(collection(db,"anggota"));ANG=s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>a.no-b.no)}return ANG}

// ---------- Navigasi ----------
const VIEWS=[
 ["ringkasan","Ringkasan",()=>can("anggota","lihat")&&can("simpanan","lihat"),()=>loadRingkasan()],
 ["anggota","Anggota",()=>can("anggota","lihat"),()=>loadAnggota()],
 ["simpanan","Simpanan",()=>can("simpanan","lihat")&&can("anggota","lihat"),()=>loadSimpanan()],
 ["bku","Buku kas",()=>can("bku","lihat"),()=>loadBku()],
 ["surat","Surat",()=>can("surat","lihat"),()=>loadSurat()],
 ["shu","SHU",()=>can("shu","lihat")&&can("anggota","lihat")&&can("simpanan","lihat"),()=>loadShu()],
 ["impor","Impor Excel",()=>can("anggota","ubah")&&can("simpanan","ubah"),()=>{}],
 ["pengguna","Pengguna",()=>me.master,()=>{resetForm();loadUsers()}],
 ["log","Log",()=>can("log","lihat"),()=>loadLog()]];
async function show(t){VIEWS.forEach(v=>$("#"+v[0]).classList.toggle("hide",v[0]!==t));
 document.querySelectorAll("#tabs button").forEach(b=>b.setAttribute("aria-selected",b.dataset.t===t));
 try{await VIEWS.find(v=>v[0]===t)[3]()}catch(e){console.error(e);const s=$("#"+t);s.insertAdjacentHTML("afterbegin",`<div class="msg err">Gagal memuat: ${esc(e.message)}</div>`)}}
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

// ---------- Ringkasan & Anggota ----------
async function loadRingkasan(){const [a,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);
 const t={pokok:0,wajib:0,sukarela:0};s.forEach(d=>{const x=d.data();t[x.jenis]+=x.jumlah});
 $("#stats").innerHTML=[[a.length,"anggota"],[rp(t.pokok),"simpanan pokok"],[rp(t.wajib),"simpanan wajib"],[rp(t.sukarela),"simpanan sukarela"],[rp(t.pokok+t.wajib+t.sukarela),"total simpanan"]].map(([b,l])=>`<div class="stat"><b>${b}</b><span>${l}</span></div>`).join("")}
async function loadAnggota(){const r=await getAnggota(true);$("#a_form").classList.toggle("hide",!can("anggota","ubah"));
 $("#ab").innerHTML=r.map(a=>`<tr><td class="n">${a.no}</td><td>${esc(a.nama)}</td><td>${esc(a.instansi)}</td><td>${esc(a.npa||"")}</td><td>${esc(a.hp||"")}</td><td>${esc(a.tglDaftar||"")}</td><td>${esc(a.status||"aktif")}</td></tr>`).join("")||`<tr><td colspan="7">Belum ada anggota. Tambahkan di atas atau gunakan Impor Excel.</td></tr>`}
$("#a_save").onclick=async()=>{const m=$("#a_msg"),nama=$("#a_nama").value.trim();if(!nama){msg(m,"Isi nama anggota.",true);return}
 try{await getAnggota();const no=Math.max(0,...ANG.map(a=>a.no||0))+1,id="A"+String(no).padStart(3,"0"),inst=$("#a_inst").value.trim();
  const b=writeBatch(db);b.set(doc(db,"anggota",id),{no,nama,instansi:inst,npa:$("#a_npa").value.trim(),hp:$("#a_hp").value.trim(),tglDaftar:today(),status:"aktif"});
  b.set(doc(db,"publik_anggota",id),{no,nama,instansi:inst,pokok:0,wajib:0,sukarela:0,total:0});await b.commit();
  await catat("tambah_anggota",nama);["a_nama","a_inst","a_npa","a_hp"].forEach(i=>$("#"+i).value="");await loadAnggota();msg(m,`Anggota nomor ${no} ditambahkan. Lanjut catat simpanan pokoknya di menu Simpanan.`)}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};

// ---------- Simpanan ----------
async function syncPublik(id){const q=await getDocs(query(collection(db,"simpanan"),where("anggotaId","==",id)));const t={pokok:0,wajib:0,sukarela:0};q.forEach(d=>{const x=d.data();t[x.jenis]+=x.jumlah});
 const a=ANG.find(x=>x.id===id);await setDoc(doc(db,"publik_anggota",id),{no:a.no,nama:a.nama,instansi:a.instansi||"",...t,total:t.pokok+t.wajib+t.sukarela},{merge:true});
 const all=await getDocs(collection(db,"simpanan"));const g={pokok:0,wajib:0,sukarela:0};all.forEach(d=>{const x=d.data();g[x.jenis]+=x.jumlah});
 await setDoc(doc(db,"publik_ringkasan","utama"),{jumlahAnggota:ANG.length,...g,diperbarui:new Date().toISOString()},{merge:true})}
async function loadSimpanan(){const [ang,s]=await Promise.all([getAnggota(true),getDocs(collection(db,"simpanan"))]);SIM=s.docs.map(d=>d.data());
 $("#s_form").classList.toggle("hide",!(can("simpanan","ubah")&&can("bku","ubah")));
 $("#s_ang").innerHTML=ang.map(a=>`<option value="${esc(a.id)}">${a.no}. ${esc(a.nama)}</option>`).join("");
 if(!$("#s_tgl").value)$("#s_tgl").value=today();if(!$("#s_per").value)$("#s_per").value=today().slice(0,7);renderSim()}
function renderSim(){const q=$("#s_cari").value.toLowerCase(),t={};SIM.forEach(x=>{(t[x.anggotaId]??={pokok:0,wajib:0,sukarela:0})[x.jenis]+=x.jumlah});
 const r=ANG.filter(a=>((a.nama||"")+(a.instansi||"")).toLowerCase().includes(q)).map(a=>{const x=t[a.id]||{pokok:0,wajib:0,sukarela:0};return {a,...x,tot:x.pokok+x.wajib+x.sukarela}});
 $("#sb").innerHTML=r.map(x=>`<tr><td class="n">${x.a.no}</td><td>${esc(x.a.nama)}</td><td>${esc(x.a.instansi)}</td><td class="n">${rp(x.pokok)}</td><td class="n">${rp(x.wajib)}</td><td class="n">${rp(x.sukarela)}</td><td class="n">${rp(x.tot)}</td></tr>`).join("");
 const s=k=>r.reduce((a,x)=>a+x[k],0);$("#sf").innerHTML=`<tr><td></td><td colspan="2">Jumlah (${r.length} anggota)</td><td class="n">${rp(s("pokok"))}</td><td class="n">${rp(s("wajib"))}</td><td class="n">${rp(s("sukarela"))}</td><td class="n">${rp(s("tot"))}</td></tr>`;
 $("#s_xls").onclick=()=>unduh(r.map(x=>({No:x.a.no,Nama:x.a.nama,Instansi:x.a.instansi,Pokok:x.pokok,Wajib:x.wajib,Sukarela:x.sukarela,Total:x.tot})),"simpanan-anggota.xlsx")}
$("#s_cari").oninput=renderSim;
$("#s_jenis").onchange=()=>{const j=$("#s_jenis").value;$("#s_perwrap").classList.toggle("hide",j!=="wajib");$("#s_jml").value=j==="pokok"?500000:j==="wajib"?50000:""};
$("#s_save").onclick=async()=>{const m=$("#s_msg"),id=$("#s_ang").value,jenis=$("#s_jenis").value,jml=+$("#s_jml").value,tgl=$("#s_tgl").value,per=$("#s_per").value;
 if(!id||!(jml>0)||!tgl){msg(m,"Lengkapi anggota, jumlah, dan tanggal.",true);return}
 if(jenis==="wajib"&&!per){msg(m,"Pilih bulan untuk simpanan wajib.",true);return}
 if(jenis==="sukarela"&&jml>15000000){msg(m,"Simpanan sukarela maksimal Rp15.000.000 per anggota (AD Pasal 34).",true);return}
 $("#s_save").disabled=true;
 try{const a=ANG.find(x=>x.id===id);const sid=jenis==="pokok"?id+"_pokok":jenis==="sukarela"?id+"_sukarela":`${id}_wajib_${per.slice(0,4)}_${per.slice(5,7)}`;
  const data={anggotaId:id,jenis,jumlah:jml,tanggal:tgl};if(jenis==="wajib")data.periode=per;
  const b=writeBatch(db);b.set(doc(db,"simpanan",sid),data);
  b.set(doc(db,"bku","sim_"+sid),{tanggal:tgl,kode:"SIM",keterangan:`Simpanan ${jenis}${jenis==="wajib"?" "+per:""} an: ${a.nama}`,kredit:jml,debit:0,ref:sid,urut:Date.now()});
  await b.commit();await syncPublik(id);await catat("setoran_simpanan",`${a.nama}: ${jenis} ${jml}`);
  const s=await getDocs(collection(db,"simpanan"));SIM=s.docs.map(d=>d.data());renderSim();msg(m,`Setoran ${jenis} ${a.nama} sebesar ${rp(jml)} tersimpan dan masuk ke buku kas.`)}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}
 $("#s_save").disabled=false};

// ---------- Buku kas ----------
async function loadBku(){const s=await getDocs(collection(db,"bku"));BKU=s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(a.tanggal||"").localeCompare(b.tanggal||"")||(a.urut||0)-(b.urut||0));
 $("#b_form").classList.toggle("hide",!can("bku","ubah"));if(!$("#b_tgl").value)$("#b_tgl").value=today();
 let sal=0,m=0,k=0;const out=[];
 $("#bb").innerHTML=BKU.map((x,i)=>{sal+=(x.kredit||0)-(x.debit||0);m+=x.kredit||0;k+=x.debit||0;out.push({No:i+1,Tanggal:x.tanggal,Kode:x.kode||"",Keterangan:x.keterangan,"Masuk (kredit)":x.kredit||0,"Keluar (debit)":x.debit||0,Saldo:sal});
  return `<tr><td class="n">${i+1}</td><td>${esc(x.tanggal)}</td><td>${esc(x.kode||"")}</td><td>${esc(x.keterangan)}</td><td class="n">${x.kredit?rp(x.kredit):""}</td><td class="n">${x.debit?rp(x.debit):""}</td><td class="n">${rp(sal)}</td></tr>`}).join("");
 $("#bf").innerHTML=`<tr><td colspan="4">Jumlah</td><td class="n">${rp(m)}</td><td class="n">${rp(k)}</td><td class="n">${rp(sal)}</td></tr>`;
 $("#b_xls").onclick=()=>unduh(out,"buku-kas-umum.xlsx")}
$("#b_save").onclick=async()=>{const m=$("#b_msg"),jml=+$("#b_jml").value,ket=$("#b_ket").value.trim(),tgl=$("#b_tgl").value;
 if(!(jml>0)||!ket||!tgl){msg(m,"Lengkapi tanggal, keterangan, dan jumlah.",true);return}
 try{const masuk=$("#b_arah").value==="masuk";await addDoc(collection(db,"bku"),{tanggal:tgl,kode:$("#b_kode").value,keterangan:ket,kredit:masuk?jml:0,debit:masuk?0:jml,urut:Date.now()});
  await catat("catat_bku",`${ket}: ${jml}`);$("#b_jml").value="";$("#b_ket").value="";await loadBku();msg(m,"Transaksi tersimpan.")}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};

// ---------- Surat ----------
const ROMAWI=["","I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII"];
async function loadSurat(){const s=await getDocs(collection(db,"surat"));SURAT=s.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")||(b.urut||0)-(a.urut||0));
 $("#r_form").classList.toggle("hide",!can("surat","ubah"));if(!$("#r_tgl").value)$("#r_tgl").value=today();nomorOtomatis();renderSurat()}
function nomorOtomatis(){const k=$("#r_jenis").value==="keluar";$("#r_pl").textContent=k?"Tujuan":"Asal surat";
 if(!k){$("#r_nomor").placeholder="Nomor dari pengirim";return}
 const tg=$("#r_tgl").value||today(),th=tg.slice(0,4),n=SURAT.filter(x=>x.jenis==="keluar"&&x.tahun===th).length+1;
 $("#r_nomor").value=`${String(n).padStart(3,"0")}/KOP.PGRI-BRU/${ROMAWI[+tg.slice(5,7)]}/${th}`}
$("#r_jenis").onchange=()=>{$("#r_nomor").value="";nomorOtomatis()};$("#r_tgl").onchange=()=>{if($("#r_jenis").value==="keluar")nomorOtomatis()};
function renderSurat(){const f=$("#r_f").value;$("#rb").innerHTML=SURAT.filter(x=>!f||x.jenis===f).map(x=>`<tr><td>${esc(x.tanggal)}</td><td>${x.jenis==="keluar"?"Keluar":"Masuk"}</td><td>${esc(x.nomor)}</td><td>${esc(x.pihak)}</td><td>${esc(x.perihal)}</td><td>${/^https?:\/\//.test(x.tautan||"")?`<a href="${esc(x.tautan)}" target="_blank" rel="noopener">Buka</a>`:""}</td></tr>`).join("")||`<tr><td colspan="6">Belum ada surat.</td></tr>`}
$("#r_f").onchange=renderSurat;
$("#r_save").onclick=async()=>{const m=$("#r_msg"),tgl=$("#r_tgl").value,hal=$("#r_hal").value.trim(),pihak=$("#r_pihak").value.trim(),nomor=$("#r_nomor").value.trim();
 if(!tgl||!hal||!pihak||!nomor){msg(m,"Lengkapi tanggal, nomor, asal/tujuan, dan perihal.",true);return}
 try{await addDoc(collection(db,"surat"),{jenis:$("#r_jenis").value,tanggal:tgl,tahun:tgl.slice(0,4),nomor,pihak,perihal:hal,tautan:$("#r_url").value.trim(),urut:Date.now()});
  await catat("catat_surat",`${nomor} - ${hal}`);["r_hal","r_pihak","r_url"].forEach(i=>$("#"+i).value="");await loadSurat();msg(m,"Surat tersimpan.")}
 catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};

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
$("#h_laba").oninput=hitungShu;$("#h_pen").oninput=hitungShu;
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
 try{await setDoc(doc(db,"shu",thn),{tahun:+thn,laba:L,persen:SHU_PCT,penerima:pen,perPenerima:pen.length?Math.round(L*SHU_PCT.pengurus/100/pen.length):0,
  hasil:SHUROWS.map(r=>({id:r.id,nama:r.nama,sukarela:r.suk,transaksi:r.tx,shuSukarela:Math.round(r.hs),shuTransaksi:Math.round(r.ht),shuAnggota:Math.round(r.ha),total:Math.round(r.z)})),disimpan:new Date().toISOString(),oleh:me.nama});
  await catat("simpan_shu",`tahun ${thn}, laba ${L}`);msg(m,`Perhitungan SHU tahun ${thn} tersimpan.`)}catch(e){msg(m,"Gagal menyimpan: "+e.message,true)}};

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
   for(const r of rows){ops.push([doc(db,"bku","B"+String(r[0]).padStart(4,"0")),{tanggal:ymd(r[1]),kode:r[2]||"",keterangan:r[3].trim(),kredit:+r[4]||0,debit:+r[5]||0,urut:r[0]}])}nB=rows.length}
  for(let i=0;i<ops.length;i+=400){const b=writeBatch(db);ops.slice(i,i+400).forEach(([ref,d])=>b.set(ref,d,{merge:true}));await b.commit();msg(m,`Menyimpan… ${Math.min(i+400,ops.length)}/${ops.length}`)}
  ANG=[];msg(m,`Selesai: ${nA} anggota dan ${nB} baris buku kas diimpor (${ops.length} catatan).`);await catat("impor_excel",`${nA} anggota, ${nB} baris BKU`)}
 catch(e){msg(m,"Impor gagal: "+e.message,true)}
 $("#imp").disabled=false};

// ---------- Log ----------
async function loadLog(){const s=await getDocs(query(collection(db,"log"),orderBy("waktu","desc"),limit(100)));
 $("#lb").innerHTML=s.docs.map(d=>{const x=d.data();return `<tr><td>${x.waktu?esc(x.waktu.toDate().toLocaleString("id-ID")):""}</td><td>${esc(x.nama)}</td><td>${esc(x.aksi)}</td><td>${esc(x.detail)}</td></tr>`}).join("")||`<tr><td colspan="4">Belum ada aktivitas.</td></tr>`}

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
 $("#ub").innerHTML=USERS.map(u=>`<tr><td>${esc(u.nama)}</td><td>${esc(u.jabatan||u.role||"")}</td><td>${esc(u.email||"")}</td><td>${u.aktif===false?"nonaktif":"aktif"}</td><td>${esc(ringkasIzin(u))}</td><td><button class="btn ghost" data-uid="${esc(u.uid)}">Ubah</button></td></tr>`).join("")}
$("#ub").onclick=e=>{const b=e.target.closest("button[data-uid]");if(!b)return;const u=USERS.find(x=>x.uid===b.dataset.uid);if(!u)return;
 editing=u;$("#ftitle").textContent="Ubah akun";$("#u_nama").value=u.nama||"";$("#u_jab").value=u.jabatan||u.role||"lainnya";$("#u_email").value=u.email||"";$("#u_email").disabled=true;
 $("#pwwrap").classList.add("hide");$("#u_aktif").checked=u.aktif!==false;$("#u_aktif").disabled=u.uid===me.uid;$("#u_master").checked=!!u.master;$("#u_master").disabled=u.uid===me.uid;buildIzin(u);syncMaster();
 $("#u_mail").classList.toggle("hide",!u.email);$("#u_msg").classList.add("hide");scrollTo({top:0,behavior:"smooth"})};
$("#u_mail").onclick=async()=>{try{await sendPasswordResetEmail(auth,editing.email);msg($("#u_msg"),"Email atur ulang kata sandi dikirim ke "+editing.email)}catch(e){msg($("#u_msg"),"Gagal mengirim: "+e.message,true)}};
$("#u_save").onclick=async()=>{const m=$("#u_msg"),nama=$("#u_nama").value.trim(),jab=$("#u_jab").value,email=$("#u_email").value.trim(),izin=collectIzin();
 if(!nama){msg(m,"Isi nama terlebih dahulu.",true);return}
 $("#u_save").disabled=true;
 try{
  if(editing){const data={nama,jabatan:jab};if(editing.uid!==me.uid){data.aktif=$("#u_aktif").checked;data.master=$("#u_master").checked;data.lihat=izin.lihat;data.ubah=izin.ubah}
   await setDoc(doc(db,"admins",editing.uid),data,{merge:true});await catat("ubah_akun",nama);await loadUsers();msg(m,"Perubahan tersimpan.")}
  else{const pw=$("#u_pw").value;
   if(!email){msg(m,"Isi email terlebih dahulu.",true);$("#u_save").disabled=false;return}
   if(pw.length<12||!/[a-z]/.test(pw)||!/[A-Z]/.test(pw)||!/\d/.test(pw)){msg(m,"Kata sandi minimal 12 karakter dan memuat huruf besar, huruf kecil, dan angka.",true);$("#u_save").disabled=false;return}
   const sec=initializeApp(firebaseConfig,"sec"+Date.now());let uid;
   try{const c=await createUserWithEmailAndPassword(getAuth(sec),email,pw);uid=c.user.uid;await signOut(getAuth(sec))}finally{await deleteApp(sec)}
   await setDoc(doc(db,"admins",uid),{nama,jabatan:jab,email,aktif:$("#u_aktif").checked,master:$("#u_master").checked,lihat:izin.lihat,ubah:izin.ubah});
   await catat("buat_akun",`${nama} (${email})`);await loadUsers();resetForm();msg(m,`Akun ${email} dibuat. Sampaikan kata sandi awal secara pribadi dan minta diganti setelah masuk.`)}}
 catch(e){msg(m,e.code==="auth/email-already-in-use"?"Email ini sudah terdaftar.":"Gagal menyimpan: "+e.message,true)}
 $("#u_save").disabled=false};
