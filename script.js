const $=id=>document.getElementById(id);
const rp=n=>"Rp "+Math.round(n).toLocaleString("id-ID");
const short=n=>n>=1e9?(n/1e9).toLocaleString("id-ID",{maximumFractionDigits:2})+" M":n>=1e6?(n/1e6).toLocaleString("id-ID",{maximumFractionDigits:1})+" jt":rp(n);
let DATA=[],metric="dana",demo=true;

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
  const keys=Object.keys(rows[0]),used=new Set(),m={};
  const rules=[["prov",/provinsi/i],["peng",/status\s*pengerjaan/i],["tipe",/status\s*institusi|tipe|jenis\s*institusi|^status$|ptn/i],["dana",/dana|pendanaan|anggaran/i],["bidang",/bidang/i],["multi",/multi|skema|durasi|lama/i],["ketua",/nama\s*ketua|ketua/i],["inst",/nama\s*institusi|institusi|perguruan|universitas/i]];
  rules.forEach(([k,re])=>{const f=keys.find(c=>!used.has(c)&&re.test(c));if(f){m[k]=f;used.add(f)}});
  const num=v=>typeof v==="number"?v:parseFloat(String(v||"0").replace(/[^\d,-]/g,"").replace(",","."))||0;
  return rows.map(r=>{
    const t=String(r[m.tipe]||"");const mv=String(r[m.multi]||"").toLowerCase();
    return{tahun:year,ketua:String(r[m.ketua]||"-").trim(),inst:String(r[m.inst]||"-").trim(),prov:String(r[m.prov]||"-").trim(),
      tipe:/ptn/i.test(t)?"PTN":/pts/i.test(t)?"PTS":(t||"Lainnya"),bidang:String(r[m.bidang]||"Lainnya").trim(),dana:num(r[m.dana]),
      multi:m.multi?(/multi|^[2-9]/.test(mv)||num(mv)>1):null,peng:m.peng?String(r[m.peng]||""):""};
  }).filter(r=>r.dana>0||r.ketua!=="-");
}
function loadFile(f){
  const rd=new FileReader();
  rd.onload=e=>{try{
    const wb=XLSX.read(e.target.result,{type:"array"});let all=[];
    wb.SheetNames.forEach(n=>{const y=(n.match(/20\d\d/)||[])[0];if(!y&&wb.SheetNames.length>1)return;
      all=all.concat(mapSheet(XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:""}),+y||2025))});
    if(!all.length)throw 0;DATA=all;demo=false;init();
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
  $("src").textContent=demo?"":"Sumber: file Excel yang diimpor";render();
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
  metric==="dana"?bars("cB",agg(R,"bidang",r=>r.dana),short):bars("cB",agg(R,"bidang",()=>1),v=>v+" judul");
  const t=agg(R,"tipe",r=>r.dana),cm={PTN:"#1b3a72",PTS:"#e8a23a"},cols=["#1b3a72","#e8a23a","#3a8f94","#999"];let a=0;
  $("ring").style.background=tot?`conic-gradient(${t.map(([_,v],i)=>{const s=a;a+=v/tot*100;return`${cm[_]||cols[i%4]} ${s}% ${a}%`}).join(",")})`:"var(--line)";
  $("lg").innerHTML=t.map(([n,v],i)=>`<div><span class="dot" style="background:${cm[n]||cols[i%4]}"></span><b>${n}</b> ${Math.round(v/tot*100)}% · ${short(v)}</div>`).join("");
  bars("cP",agg(R,"prov",r=>r.dana).slice(0,5),short);
  const top=[...R].sort((a,b)=>b.dana-a.dana).slice(0,5);
  $("tb").innerHTML=top.length?top.map((r,i)=>`<tr><td>${i+1}</td><td>${r.ketua}</td><td>${r.inst}</td><td><span class="tag">${r.bidang}</span></td><td>${r.tahun}</td><td class="r">${rp(r.dana)}</td></tr>`).join(""):'<tr><td colspan="6" class="empty">Tidak ada data</td></tr>';
  const s26=R.filter(r=>r.peng);
  $("note").textContent=s26.length?"Status Pengerjaan (2026): "+agg(s26,"peng",()=>1).map(([n,v])=>n+" "+v).join(" · "):"";
  window._R=R;
}
Object.keys(F).forEach(id=>$(id).onchange=e=>{F[id]=e.target.value;render()});
$("tg").onclick=e=>{const m=e.target.dataset.m;if(!m)return;metric=m;[...$("tg").children].forEach(b=>b.classList.toggle("on",b.dataset.m===m));render()};
$("rst").onclick=init;
$("imp").onclick=()=>$("file").click();
$("file").onchange=e=>e.target.files[0]&&loadFile(e.target.files[0]);
$("exp").onclick=()=>{const R=window._R||[];const h=["tahun","ketua","inst","bidang","tipe","prov","dana"];
  const csv=[h.join(",")].concat(R.map(r=>h.map(k=>'"'+String(r[k]).replace(/"/g,'""')+'"').join(","))).join("\n");
  const a=document.createElement("a");a.href="data:text/csv;charset=utf-8,"+encodeURIComponent(csv);a.download="rikub-filtered.csv";a.click()};
DATA=sample();init();