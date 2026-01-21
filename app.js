function readFile(file, sheetName=null) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, {type:file.name.endsWith(".csv")?"string":"binary"});
      const sh = file.name.endsWith(".csv")
        ? wb.Sheets[wb.SheetNames[0]]
        : wb.Sheets[sheetName||wb.SheetNames[0]];
      if(!sh) return rej("Sheet missing");
      res(XLSX.utils.sheet_to_json(sh,{defval:""}));
    };
    file.name.endsWith(".csv") ? r.readAsText(file) : r.readAsBinaryString(file);
  });
}

function normSku(raw){
  if(!raw) return null;
  raw = raw.toString().trim().toUpperCase().replace(/\s+/g,"");
  let style,size,disp;
  if(raw.includes("-")){
    [style,size]=raw.split("-");
    disp=size;
  } else {
    style=raw; size="FS"; disp="";
  }
  return {canon:`${style}-${size}`,style,size,disp};
}

async function generate(){
  const sales = await readFile(saleFile.files[0]);
  const uniware = await readFile(uniwareFile.files[0]);
  const inward = await readFile(binFile.files[0],"Inward");

  const stock={}, master=[];
  uniware.forEach(r=>{
    const s=normSku(r["Sku Code"]);
    const q=+r["Available (ATP)"];
    if(!s||q<=0) return;
    stock[s.canon]=stock[s.canon]||{};
    stock[s.canon].godown=(stock[s.canon].godown||0)+q;
  });

  inward.forEach(r=>{
    const s=normSku(r["SKU"]);
    const b=r["Bin"]; const q=+r["Qty"];
    if(!s||!b||q<=0) return;
    stock[s.canon]=stock[s.canon]||{};
    stock[s.canon].godown=Math.max(0,(stock[s.canon].godown||0)-q);
    stock[s.canon][b]=(stock[s.canon][b]||0)+q;
  });

  for(const sku in stock){
    for(const b in stock[sku]){
      if(stock[sku][b]>0)
        step1.innerHTML+=`<tr><td>${sku}</td><td>${b}</td><td>${stock[sku][b]}</td></tr>`;
    }
  }

  const demand={}, invalid=[];
  sales.forEach(r=>{
    const mp=r["Channel Name"]||"";
    const s=normSku(r["Item SKU Code"]);
    if(!s){ invalid.push({sku:"",mp,reason:"Invalid SKU"}); return;}
    demand[mp]=demand[mp]||{};
    demand[mp][s.canon]=(demand[mp][s.canon]||0)+1;
  });

  for(const mp in demand){
    for(const sku in demand[mp]){
      const total=Object.values(stock[sku]||{}).reduce((a,b)=>a+b,0);
      step2.innerHTML+=`<tr><td>${sku}</td><td>${mp}</td><td>${demand[mp][sku]}</td><td>${total}</td></tr>`;
      step4.innerHTML+=`<tr><td>${mp}</td><td>${sku}</td><td>${demand[mp][sku]}</td><td>${total}</td></tr>`;
    }
  }

  invalid.forEach(r=>{
    step3.innerHTML+=`<tr><td>${r.sku}</td><td>${r.mp}</td><td>${r.reason}</td></tr>`;
  });

  for(const mp in demand){
    for(const sku in demand[mp]){
      let need=demand[mp][sku];
      const bins=Object.entries(stock[sku]||{}).sort((a,b)=>a[0]=="godown"?-1:1);
      const {style,disp}=normSku(sku);
      for(const [b,q] of bins){
        if(need<=0) break;
        const pick=Math.min(q,need);
        step5.innerHTML+=`<tr><td>${sku}</td><td>${style}</td><td>${disp}</td><td>${b}</td><td>${pick}</td></tr>`;
        need-=pick;
      }
      if(need>0)
        shortage.innerHTML+=`<tr><td>${sku}</td><td>${mp}</td><td>${need}</td></tr>`;
    }
  }
}
