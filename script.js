const $=id=>document.getElementById(id);
const rp=n=>"Rp "+Math.round(n).toLocaleString("id-ID");
const short=n=>n>=1e9?(n/1e9).toLocaleString("id-ID",{maximumFractionDigits:2})+" M":n>=1e6?(n/1e6).toLocaleString("id-ID",{maximumFractionDigits:1})+" jt":rp(n);
let DATA=[],MT={B:"dana",T:"n",P:"dana",M:"dana"},demo=true,META=null,LASTMAP={};

/* ---------- Data contoh (dipakai sampai file Excel diimpor) ---------- */
function sample(){
  let s=7;const r=()=>(s=s*16807%2147483647)/2147483647;
  const B=["Pangan","Kesehatan","Energi","Produk Rekayasa Keteknikan","Sosial Humaniora","Maritim"];
  const I=[["Universitas Gadjah Mada","DI Yogyakarta","PTN"],["Institut Pertanian Bogor","Jawa Barat","PTN"],["Institut Teknologi Bandung","Jawa Barat","PTN"],["Universitas Diponegoro","Jawa Tengah","PTN"],["Universitas Udayana","Bali","PTN"],["Universitas Amikom Yogyakarta","DI Yogyakarta","PTS"],["Universitas Telkom","Jawa Barat","PTS"],["Universitas Muhammadiyah Malang","Jawa Timur","PTS"],["Universitas Airlangga","Jawa Timur","PTN"],["Universitas Sultan Ageng Tirtayasa","Banten","PTN"],["Universitas Kristen Satya Wacana","Jawa Tengah","PTS"],["Universitas Mahasaraswati Denpasar","Bali","PTS"]];
  const N=["Budi Santoso","Siti Nurhaliza","Kuwat Triyana","Ratna Dewi","Agus Prasetyo","Dewi Lestari","Hendra Wijaya","Maya Sari","Rizal Fauzi","Putri Anggraini"];
  const out=[];
  for(let i=0;i<68;i++){const y=i<30?2025:2026,ins=I[Math.floor(r()*I.length)];
    out.push({tahun:y,ketua:N[Math.floor(r()*N.length)]+" "+(i+1),inst:ins[0],prov:ins[1],tipe:ins[2],bidang:B[Math.floor(r()*B.length)],
      dana:Math.round((3e8+r()*1.6e9)/1000)*1000,multi:r()<.85,peng:y==2026?["Berjalan","Selesai","Persiapan"][Math.floor(r()*3)]:""});}
  return out;
}

/* ---------- Konsolidasi & standardisasi kolom (2025 vs 2026) ---------- */
function mapSheet(rows,year){
  if(!rows.length)return[];
  const keys=Object.keys(rows[0]),used=new Set(),m={},nk=c=>String(c).replace(/\s+/g," ").trim();
  const EXI=/provinsi|kabupaten|kota|tipe|status|jenis/i,EXK=/provinsi|kabupaten|kota|tipe|status|jenis|institusi|perguruan/i;
  const rules=[["prov",null,/provinsi/i],["peng",null,/status\s*pengerjaan/i],["tipe",null,/^status\s*ins\w*$/i,/status\s*ins\w*/i],["dana",null,/dana|pendanaan|anggaran/i],["bidang",null,/bidang/i],["multi",null,/multi|skema|durasi|lama/i],["ketua",EXK,/^nama$/i,/^nama\s*ketua/i,/ketua/i],["inst",EXI,/nama\s*institusi/i,/institusi.*ketua|perguruan/i,/institusi|universitas/i]];
  rules.forEach(([k,ex,...res])=>{for(const re of res){const f=keys.find(c=>{const n=nk(c);return !used.has(c)&&re.test(n)&&!(ex&&ex.test(n))});if(f){m[k]=f;used.add(f);break}}});const isT=v=>{v=String(v||"");return v.length<=30&&/pt[ns]|negeri|swasta/i.test(v)},sc=c=>rows.filter(r=>isT(r[c])).length/rows.length;
  if(!m.tipe||sc(m.tipe)<.3){const cand=keys.filter(c=>!used.has(c)).map(c=>[c,sc(c)]).filter(x=>x[1]>=.3).sort((a,b)=>b[1]-a[1]);if(cand.length){m.tipe=cand[0][0];used.add(m.tipe)}}
  LASTMAP=m;
  const num=v=>typeof v==="number"?v:parseFloat(String(v||"0").replace(/[^\d,-]/g,"").replace(",","."))||0;
  return rows.map(r=>{
    const t=String(r[m.tipe]||"");const mv=String(r[m.multi]||"").toLowerCase();
    return{tahun:year,ketua:String(r[m.ketua]||"-").replace(/\s+/g," ").trim(),inst:String(r[m.inst]||"-").trim(),prov:String(r[m.prov]||"-").trim(),
      tipe:/ptn|negeri/i.test(t)?"PTN":/pts|swasta/i.test(t)?"PTS":(t||"Lainnya"),bidang:String(r[m.bidang]||"Lainnya").trim(),dana:num(r[m.dana]),
      multi:m.multi?(/multi|^[2-9]/.test(mv)||num(mv)>1):null,peng:m.peng?String(r[m.peng]||""):""};
  }).filter(r=>r.dana>0||r.ketua!=="-");
}
function loadFile(f){
  const rd=new FileReader();
  rd.onload=e=>{try{
    const wb=XLSX.read(e.target.result,{type:"array"});let all=[],log=[];
    wb.SheetNames.forEach(n=>{const y=(n.match(/20\d\d/)||[])[0];if(!y&&wb.SheetNames.length>1)return;
      all=all.concat(mapSheet(XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:""}),+y||2025));log.push(n+" → "+Object.entries(LASTMAP).map(([k,v])=>k+": "+v).join(" | "))});
    if(!all.length)throw 0;DATA=all;demo=false;META={name:f.name,time:Date.now(),map:log.join("\n")};persist().then(ok=>{if(ok!==false)return init();loadSaved().then(h=>{if(!h){DATA=sample();META=null;demo=true}init()})});
  }catch(x){alert("File tidak terbaca. Pastikan ada sheet bernama 2025 dan/atau 2026 dengan baris header di baris pertama.")}};
  rd.readAsArrayBuffer(f);
}

/* ---------- Filter & render ---------- */
const F={fY:"Semua",fB:"Semua",fT:"Semua",fP:"Semua"},FK={fY:"tahun",fB:"bidang",fT:"tipe",fP:"prov"};
function fill(id,vals){const el=$(id);el.innerHTML=["Semua",...vals].map(v=>`<option>${v}</option>`).join("");el.value=F[id]}
function init(){
  F.fY=F.fB=F.fT=F.fP="Semua";
  const u=k=>[...new Set(DATA.map(d=>d[k]))].sort();
  fill("fY",u("tahun"));fill("fB",u("bidang"));fill("fT",u("tipe"));fill("fP",u("prov"));
  $("src").textContent=demo?"":META?"Data tersimpan: "+META.name+" · "+new Date(META.time).toLocaleString("id-ID"):"Sumber: file Excel yang diimpor";$("src").title=META?META.map:"";ui();render();
}
function agg(rows,k,fn){const m={};rows.forEach(r=>{m[r[k]]=(m[r[k]]||0)+fn(r)});return Object.entries(m).sort((a,b)=>b[1]-a[1])}
function bars(el,list,fmt){const mx=Math.max(1,...list.map(x=>x[1]));
  $(el).innerHTML=list.length?list.map(([n,v])=>`<div class="row"><div class="n" title="${n}">${n}</div><div class="trk"><div class="fill" style="width:${v/mx*100}%"></div></div><div>${fmt(v)}</div></div>`).join(""):'<div class="empty">Tidak ada data</div>'}
function render(){
  const R=DATA.filter(d=>Object.keys(FK).every(k=>F[k]==="Semua"||String(d[FK[k]])===F[k]));
  const tot=R.reduce((a,b)=>a+b.dana,0),mv=R.filter(r=>r.multi!==null),mm=mv.filter(r=>r.multi).length;
  $("k1").textContent=rp(tot);$("k1s").textContent=R.length?"rata-rata "+short(tot/R.length)+" / judul":"";
  $("k2").textContent=R.length+" Judul";$("k2s").textContent=[2025,2026].map(y=>y+": "+R.filter(r=>r.tahun===y).length).join(" · ");
  $("k3").textContent=new Set(R.map(r=>r.inst)).size+" Institusi";
  if(mv.length){const p=Math.round(mm/mv.length*100);$("k4").textContent=p+"% Multi-Years";$("k4s").textContent=(100-p)+"% Single-Year"}else{$("k4").textContent="–";$("k4s").textContent="kolom skema tidak ditemukan"}
  MT.B==="dana"?bars("cB",agg(R,"bidang",r=>r.dana),short):bars("cB",agg(R,"bidang",()=>1),v=>v+" judul");
  const vT=MT.T==="dana"?r=>r.dana:()=>1,t=agg(R,"tipe",vT),tt=t.reduce((x,y)=>x+y[1],0),fT=v=>MT.T==="dana"?short(v):v+" judul",cm={PTN:"#1b3a72",PTS:"#e8a23a"},cols=["#1b3a72","#e8a23a","#3a8f94","#999"];let a=0;
  $("hT").textContent="Status Institusi (porsi "+(MT.T==="dana"?"dana":"jumlah judul")+")";$("ring").style.background=tt?`conic-gradient(${t.map(([_,v],i)=>{const s=a;a+=v/tt*100;return`${cm[_]||cols[i%4]} ${s}% ${a}%`}).join(",")})`:"var(--line)";
  $("lg").innerHTML=t.map(([n,v],i)=>`<div><span class="dot" style="background:${cm[n]||cols[i%4]}"></span><b>${n}</b> ${Math.round(v/tt*100)}%<div style="font-size:12px;color:var(--mute);margin-left:16px">${R.filter(r=>r.tipe===n).length} judul · ${short(R.filter(r=>r.tipe===n).reduce((x,y)=>x+y.dana,0))}</div></div>`).join("");
  let b0=0;$("ring").innerHTML=tt?t.map(([n,v])=>{const p=v/tt*100,md=(b0+p/2)/100*2*Math.PI;b0+=p;if(p<6)return"";return`<span style="position:absolute;left:50%;top:50%;z-index:1;transform:translate(-50%,-50%) translate(${Math.sin(md)*54}px,${-Math.cos(md)*54}px);color:#fff;font:700 11px/1.15 system-ui,sans-serif;text-align:center;text-shadow:0 1px 2px rgba(0,0,0,.35)">${n}<br>${Math.round(p)}%</span>`}).join(""):"";
  bars("cP",agg(R,"prov",MT.P==="dana"?r=>r.dana:()=>1).slice(0,5),MT.P==="dana"?short:v=>v+" judul");
  const top=[...R].sort((a,b)=>b.dana-a.dana).slice(0,10);
  $("tb").innerHTML=top.length?top.map((r,i)=>`<tr><td>${i+1}</td><td>${r.ketua}</td><td>${r.inst}</td><td><span class="tag">${r.bidang}</span></td><td>${r.tahun}</td><td class="r">${rp(r.dana)}</td></tr>`).join(""):'<tr><td colspan="6" class="empty">Tidak ada data</td></tr>';
  $("cnt").textContent=" · "+top.length+" dari "+R.length+" riset";
  const s26=R.filter(r=>r.peng);
  $("note").textContent=s26.length?"Status Pengerjaan (2026): "+agg(s26,"peng",()=>1).map(([n,v])=>n+" "+v).join(" · "):"";
  drawMap(R);
  window._R=R;
}
Object.keys(F).forEach(id=>$(id).onchange=e=>{F[id]=e.target.value;render()});
document.querySelectorAll(".tg").forEach(g=>g.onclick=e=>{const m=e.target.dataset.m;if(!m)return;MT[g.dataset.c]=m;[...g.children].forEach(b=>b.classList.toggle("on",b.dataset.m===m));render()});
$("imp").onclick=()=>$("file").click();
$("file").onchange=e=>e.target.files[0]&&loadFile(e.target.files[0]);
$("exp").onclick=()=>{const R=window._R||[];const h=["tahun","ketua","inst","bidang","tipe","prov","dana"];
  const csv=[h.join(",")].concat(R.map(r=>h.map(k=>'"'+String(r[k]).replace(/"/g,'""')+'"').join(","))).join("\n");
  const a=document.createElement("a");a.href="data:text/csv;charset=utf-8,"+encodeURIComponent(csv);a.download="rikub-filtered.csv";a.click()};
/* ---------- Peta Indonesia (makin gelap = makin banyak) ---------- */
const PAL={diaceh:"aceh",nanggroeacehdarussalam:"aceh",daerahistimewaaceh:"aceh",bangkabelitung:"kepulauanbangkabelitung",babel:"kepulauanbangkabelitung",diyogyakarta:"daerahistimewayogyakarta",yogyakarta:"daerahistimewayogyakarta",diy:"daerahistimewayogyakarta",jakarta:"dkijakarta",dki:"dkijakarta",ntb:"nusatenggarabarat",ntt:"nusatenggaratimur",kepri:"kepulauanriau"};
const pkey=s=>{const k=String(s||"").toLowerCase().replace(/[^a-z]/g,"").replace(/^provinsi/,"");return PAL[k]||k};
let MAPREADY=false;
function drawMap(R){
  const M=window.RIKUB_MAP,svg=$("peta");if(!M||!svg)return;
  if(!MAPREADY){
    svg.setAttribute("viewBox",M.viewBox);
    svg.innerHTML=M.prov.map(p=>`<path data-k="${p.k}" data-n="${p.n}" d="${p.d}"/>`).join("");
    const tip=$("tip");
    svg.onmousemove=e=>{const pa=e.target.closest("path");if(!pa){tip.hidden=true;return}
      const o=(window._ST||{})[pa.dataset.k],r=svg.parentNode.getBoundingClientRect();
      tip.innerHTML="<b>"+pa.dataset.n+"</b><br>"+(o?o.n+" judul<br>"+rp(o.d):"Belum ada data");
      tip.hidden=false;tip.style.left=Math.min(e.clientX-r.left+14,r.width-170)+"px";tip.style.top=(e.clientY-r.top+14)+"px"};
    svg.onmouseleave=()=>{tip.hidden=true};
    MAPREADY=true}
  const st={};R.forEach(r=>{const k=pkey(r.prov),o=st[k]||(st[k]={n:0,d:0,name:r.prov});o.n++;o.d+=r.dana});window._ST=st;
  const val=o=>MT.M==="dana"?o.d:o.n,vs=Object.values(st).map(val),mx=vs.length?Math.max(...vs):0,mn=vs.length?Math.min(...vs):0,a=[190,215,240],b=[27,58,114];
  const keys=new Set(M.prov.map(p=>p.k));
  svg.querySelectorAll("path").forEach(pa=>{const o=st[pa.dataset.k];
    if(!o){pa.style.fill="var(--line)";return}
    const t=mx>mn?(val(o)-mn)/(mx-mn):1;pa.style.fill="rgb("+a.map((x,i)=>Math.round(x+(b[i]-x)*t))+")"});
  const f=v=>MT.M==="dana"?short(v):v+" judul";$("lmin").textContent=vs.length?"Sedikit · "+f(mn):"";$("lmax").textContent=vs.length?f(mx)+" · Banyak":"";
  const miss=Object.entries(st).filter(([k])=>!keys.has(k)).map(([,o])=>o.name);
  $("mnote").textContent=(vs.length?Object.keys(st).length-miss.length+" provinsi memiliki data. Abu-abu = belum ada data. ":"Tidak ada data. ")+(miss.length?"Tidak ditemukan di peta: "+miss.join(", ")+".":"")};

/* ---------- Penyimpanan data: Supabase (login admin) atau browser ---------- */
const KEY="rikub:data:v1",CFG=window.RIKUB_CFG||{},
SB=(CFG.SUPABASE_URL&&CFG.SUPABASE_ANON_KEY&&window.supabase)?window.supabase.createClient(CFG.SUPABASE_URL,CFG.SUPABASE_ANON_KEY):null;
let ADMIN=false;
function ui(){$("imp").hidden=!!SB&&!ADMIN;$("login").hidden=!SB||ADMIN;$("logout").hidden=!SB||!ADMIN;$("clr").hidden=!(META&&(!SB||ADMIN))}
async function persist(){
  if(SB){const {error}=await SB.from("rikub_snapshot").upsert({id:1,meta:META,data:DATA,updated_at:new Date().toISOString()});
    if(error){alert("Data gagal disimpan ke server. Pastikan Anda login sebagai admin yang terdaftar.\n"+error.message);return false}return true}
  try{localStorage.setItem(KEY,JSON.stringify({meta:META,data:DATA}))}catch(x){alert("Data tampil, tetapi gagal disimpan permanen (penyimpanan browser penuh atau diblokir).")}return true}
async function loadSaved(){
  try{
    if(SB){const {data,error}=await SB.from("rikub_snapshot").select("meta,data").eq("id",1).maybeSingle();
      if(!error&&data&&data.data&&data.data.length){DATA=data.data;META=data.meta;demo=false;return true}return false}
    const o=JSON.parse(localStorage.getItem(KEY)||"null");if(o&&o.data&&o.data.length){DATA=o.data;META=o.meta;demo=false;return true}
  }catch(x){}return false}
async function checkAdmin(){ADMIN=false;if(SB){try{const {data:{session}}=await SB.auth.getSession();if(session){const {data}=await SB.from("admins").select("email").limit(1);ADMIN=!!(data&&data.length)}}catch(x){}}ui()}
$("clr").onclick=async()=>{if(!confirm("Hapus data yang tersimpan dan kembali ke data contoh?"))return;
  if(SB){const {data,error}=await SB.from("rikub_snapshot").delete().eq("id",1).select();if(error||!data||!data.length){alert("Gagal menghapus. Pastikan Anda login sebagai admin yang terdaftar.");return}}
  else{try{localStorage.removeItem(KEY)}catch(x){}}
  META=null;demo=true;DATA=sample();init()};
if(SB){
  $("login").onclick=()=>{$("lerr").textContent="";$("dlg").showModal()};
  $("lcancel").onclick=()=>$("dlg").close();
  $("logout").onclick=async()=>{await SB.auth.signOut();checkAdmin()};
  let gagal=0,kunci=0;
  $("lf").onsubmit=async e=>{e.preventDefault();
    if(Date.now()<kunci){$("lerr").textContent="Terlalu banyak percobaan. Coba lagi sebentar.";return}
    if(!CFG.ADMIN_EMAIL){$("lerr").textContent="ADMIN_EMAIL belum diisi di config.js.";return}
    $("lerr").textContent="Memproses…";
    const {error}=await SB.auth.signInWithPassword({email:CFG.ADMIN_EMAIL,password:$("pw").value});
    if(error){gagal++;if(gagal>=5){kunci=Date.now()+30000;gagal=0}$("lerr").textContent="PIN salah.";$("pw").value="";return}
    gagal=0;await checkAdmin();
    if(!ADMIN){$("lerr").textContent="Akun admin belum terdaftar di tabel admins Supabase.";await SB.auth.signOut();return}
    $("pw").value="";$("dlg").close()};
  SB.auth.onAuthStateChange(()=>setTimeout(checkAdmin,0));
}
(async()=>{if(!(await loadSaved()))DATA=sample();init();await checkAdmin()})();

/* ---------- Menu tab di atas ---------- */
document.querySelectorAll("#menu button").forEach(b=>b.onclick=()=>{
  document.querySelectorAll("#menu button").forEach(x=>x.classList.toggle("on",x===b));
  document.querySelectorAll(".pane").forEach(p=>p.classList.toggle("on",p.dataset.t===b.dataset.t));
});