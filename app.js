let mpData = {};
let activeMP = "";
let filteredPick = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

function readXlsx(file, sheetName=null) {
  return new Promise(res => {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const sheet = sheetName ? wb.Sheets[sheetName] : wb.Sheets[wb.SheetNames[0]];
      res(XLSX.utils.sheet_to_json(sheet));
    };
    r.readAsBinaryString(file);
  });
}

function parseSku(sku) {
  if (!sku.includes("-")) return { style: sku, size: "FS" };
  const [s, z] = sku.split("-");
  return { style: s, size: z || "FS" };
}

async function generate() {
  const sale = await readXlsx(saleFile.files[0]);
  const uni = await readXlsx(uniwareFile.files[0]);
  const bin = await readXlsx(binFile.files[0], "Inward");

  /* DEMAND MP → SKU */
  const demand = {};
  sale.forEach(r => {
    const sku = r["Item SKU Code (Sku Code)"];
    const mp = r["Channel Name (MP)"] || "UNKNOWN";
    if (!sku) return;
    demand[mp] = demand[mp] || {};
    demand[mp][sku] = (demand[mp][sku] || 0) + 1;
  });

  /* MASTER STOCK */
  const stock = {};

  // Uniware stock (+)
  uni.forEach(r => {
    const sku = r["Sku Code"];
    const binId = r["Shelf"];
    const qty = Number(r["Available (ATP)"]);
    if (!sku || !binId || !qty) return;
    stock[sku] = stock[sku] || {};
    stock[sku][binId] = (stock[sku][binId] || 0) + qty;
  });

  // Bin inward stock (-)
  bin.forEach(r => {
    const sku = r["SKU"];
    const binId = r["Bin"];
    const qty = Number(r["Qty"]);
    if (!sku || !binId || !qty) return;
    stock[sku] = stock[sku] || {};
    stock[sku][binId] = (stock[sku][binId] || 0) - qty;
  });

  /* PICK GENERATION */
  mpData = {};

  for (let mp in demand) {
    mpData[mp] = [];

    for (let sku in demand[mp]) {
      let need = demand[mp][sku];
      const bins = Object.entries(stock[sku] || {})
        .filter(([_, q]) => q > 0)
        .sort((a,b)=>{
          if(a[0].toLowerCase()==="godown") return -1;
          if(b[0].toLowerCase()==="godown") return 1;
          return a[0].localeCompare(b[0]);
        });

      const { style, size } = parseSku(sku);

      for (let [binId, qty] of bins) {
        if (need <= 0) break;
        const pick = Math.min(qty, need);
        mpData[mp].push({
          SKU: sku,
          Style: style,
          Size: size,
          BIN: binId,
          Unit: pick
        });
        need -= pick;
      }
    }

    mpData[mp].sort((a,b)=>{
      if(a.SKU!==b.SKU) return a.SKU.localeCompare(b.SKU);
      return SIZE_ORDER.indexOf(a.Size) - SIZE_ORDER.indexOf(b.Size);
    });
  }

  buildTabs();
}

function buildTabs() {
  mpTabs.innerHTML = "";
  const mps = Object.keys(mpData);
  if (!mps.length) return;

  mps.forEach((mp,i)=>{
    const t = document.createElement("div");
    t.className = "tab" + (i===0 ? " active":"");
    t.innerText = mp;
    t.onclick = ()=>switchMP(mp);
    mpTabs.appendChild(t);
  });

  switchMP(mps[0]);
}

function switchMP(mp) {
  activeMP = mp;
  [...mpTabs.children].forEach(t=>{
    t.classList.toggle("active", t.innerText===mp);
  });
  applyFilters();
}

function applyFilters() {
  const s = fSku.value.toLowerCase();
  const st = fStyle.value.toLowerCase();
  const sz = fSize.value.toLowerCase();
  const b = fBin.value.toLowerCase();

  filteredPick = (mpData[activeMP] || []).filter(r =>
    r.SKU.toLowerCase().includes(s) &&
    r.Style.toLowerCase().includes(st) &&
    r.Size.toLowerCase().includes(sz) &&
    r.BIN.toLowerCase().includes(b)
  );
  render();
}

function render() {
  tbody.innerHTML = "";
  filteredPick.forEach(r=>{
    tbody.innerHTML += `
      <tr>
        <td>${r.SKU}</td>
        <td>${r.Style}</td>
        <td>${r.Size}</td>
        <td>${r.BIN}</td>
        <td>${r.Unit}</td>
      </tr>`;
  });
}

function exportExcel() {
  const wb = XLSX.utils.book_new();
  for (let mp in mpData) {
    const ws = XLSX.utils.json_to_sheet(mpData[mp]);
    XLSX.utils.book_append_sheet(wb, ws, mp.substring(0,31));
  }
  XLSX.writeFile(wb, "MP_Wise_Pick_List.xlsx");
}
