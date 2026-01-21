let mpData = {};
let shortageData = [];
let activeTab = "";
let currentView = [];

const SIZE_ORDER = ["FS","XS","S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];

function readFile(file, sheetName=null) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, {type:file.name.endsWith(".csv")?"string":"binary"});
      const sh = file.name.endsWith(".csv")
        ? wb.Sheets[wb.SheetNames[0]]
        : wb.Sheets[sheetName || wb.SheetNames[0]];
      if (!sh) return rej("Required sheet missing");
      res(XLSX.utils.sheet_to_json(sh,{defval:""}));
    };
    file.name.endsWith(".csv") ? r.readAsText(file) : r.readAsBinaryString(file);
  });
}

function normSku(raw){
  if(!raw) return null;
  raw = raw.toString().trim().toUpperCase().replace(/\s+/g,"");
  let style,size,disp;
  if(raw.includes("-")) { [style,size]=raw.split("-"); disp=size; }
  else { style=raw; size="FS"; disp=""; }
  if(!style) return null;
  return {canon:`${style}-${size}`,style,disp};
}

async function generate(){
  document.getElementById("msg").innerText = "Processing...";
  mpData = {}; shortageData = {}; currentView = [];

  const sales = await readFile(saleFile.files[0]);
  const uniware = await readFile(uniwareFile.files[0]);
  const inward = await readFile(binFile.files[0],"Inward");

  // ---- MASTER STOCK ----
  const stock = {};
  for (const r of uniware) {
    const s = normSku(r["Sku Code"]);
    const q = +r["Available (ATP)"];
    if(!s||q<=0) continue;
    stock[s.canon] = stock[s.canon] || {};
    stock[s.canon].godown = (stock[s.canon].godown||0)+q;
  }
  for (const r of inward) {
    const s = normSku(r["SKU"]);
    const b = r["Bin"];
    const q = +r["Qty"];
    if(!s||!b||q<=0) continue;
    stock[s.canon] = stock[s.canon] || {};
    stock[s.canon].godown = Math.max(0,(stock[s.canon].godown||0)-q);
    stock[s.canon][b] = (stock[s.canon][b]||0)+q;
  }

  // ---- DEMAND ----
  const demand = {};
  for (const r of sales) {
    const s = normSku(r["Item SKU Code"]);
    const mp = r["Channel Name"] || "UNKNOWN";
    if(!s){ shortageData[mp]=(shortageData[mp]||0)+1; continue; }
    demand[mp]=demand[mp]||{};
    demand[mp][s.canon]=(demand[mp][s.canon]||0)+1;
  }

  // ---- PICKING ----
  for (const mp in demand) {
    mpData[mp]=[];
    for (const sku in demand[mp]) {
      let need=demand[mp][sku];
      const bins = Object.entries(stock[sku]||{}).filter(x=>x[1]>0)
        .sort((a,b)=>a[0]=="godown"?-1:b[0]=="godown"?1:a[0].localeCompare(b[0]));
      const info = normSku(sku);
      for (const [bin,q] of bins) {
        if(need<=0) break;
        const pick=Math.min(q,need);
        mpData[mp].push({SKU:sku,Style:info.style,Size:info.disp,BIN:bin,Units:pick});
        need-=pick;
      }
      if(need>0){
        shortageData[mp]= (shortageData[mp]||0)+need;
      }
    }
  }

  buildTabs();
  document.getElementById("msg").innerText = "✅ Report generated";
}

function buildTabs(){
  const t=document.getElementById("tabs");
  t.innerHTML="";
  const mps=Object.keys(mpData);
  mps.forEach((mp,i)=>{
    const d=document.createElement("div");
    d.className="tab"+(i==0?" active":"");
    d.innerText=mp;
    d.onclick=()=>switchTab(mp);
    t.appendChild(d);
  });
  if(Object.keys(shortageData).length){
    const d=document.createElement("div");
    d.className="tab";
    d.innerText="SHORTAGE";
    d.onclick=()=>switchTab("SHORTAGE");
    t.appendChild(d);
  }
  switchTab(mps[0]||"SHORTAGE");
}

function switchTab(tab){
  activeTab=tab;
  document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("active",t.innerText==tab));
  currentView = tab=="SHORTAGE"
    ? Object.entries(shortageData).map(([mp,u])=>({SKU:"",Style:"",Size:"",BIN:mp,Units:u}))
    : mpData[tab]||[];
  applySort();
}

function applySort(){
  const k=document.getElementById("sortBy").value;
  if(k=="Size") currentView.sort((a,b)=>SIZE_ORDER.indexOf(a.Size||"FS")-SIZE_ORDER.indexOf(b.Size||"FS"));
  else if(k) currentView.sort((a,b)=>(a[k]||"").localeCompare(b[k]||""));
  render();
}

function render(){
  const tb=document.getElementById("tbody");
  let h="";
  for(const r of currentView){
    h+=`<tr><td>${r.SKU}</td><td>${r.Style}</td><td>${r.Size}</td><td>${r.BIN}</td><td>${r.Units}</td></tr>`;
  }
  tb.innerHTML=h;
}

function exportExcel(){
  const wb=XLSX.utils.book_new();
  for(const mp in mpData){
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(mpData[mp]),mp);
  }
  if(Object.keys(shortageData).length){
    const sh=Object.entries(shortageData).map(([MP,Units])=>({MP,Units}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sh),"SHORTAGE");
  }
  XLSX.writeFile(wb,"Pick_List_MP_Wise.xlsx");
}
