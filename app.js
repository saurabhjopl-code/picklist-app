let mpData = {};
let activeMP = "";
let currentView = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

/* ---------- FILE READ (SAFE) ---------- */
function readFile(file, sheetName=null) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      const isCSV = file.name.toLowerCase().endsWith(".csv");
      const wb = XLSX.read(
        e.target.result,
        { type: isCSV ? "string" : "binary" }
      );

      let sheet;
      if (isCSV) {
        sheet = wb.Sheets[wb.SheetNames[0]];
      } else {
        sheet = sheetName
          ? wb.Sheets[sheetName]
          : wb.Sheets[wb.SheetNames[0]];
      }

      if (!sheet) {
        rej(`Sheet "${sheetName}" not found in ${file.name}`);
        return;
      }

      res(XLSX.utils.sheet_to_json(sheet));
    };

    isCSV ? reader.readAsText(file) : reader.readAsBinaryString(file);
  });
}

/* ---------- SKU ---------- */
function parseSku(sku) {
  if (!sku.includes("-")) return { style: sku, size: "FS" };
  const [s, z] = sku.split("-");
  return { style: s, size: z || "FS" };
}

/* ---------- UPLOAD STATUS ---------- */
["saleFile","uniwareFile","binFile"].forEach(id => {
  document.getElementById(id).addEventListener("change", e => {
    document.getElementById(id.replace("File","Status")).innerText =
      e.target.files.length ? "✔ Uploaded" : "";
  });
});

/* ---------- GENERATE ---------- */
async function generate() {
  msg.innerText = "";

  try {
    const sale = await readFile(saleFile.files[0]);
    const uni = await readFile(uniwareFile.files[0]);
    const bin = await readFile(binFile.files[0], "Inward");

    /* DEMAND */
    const demand = {};
    sale.forEach(r => {
      const sku = r["Item SKU Code (Sku Code)"];
      const mp = r["Channel Name (MP)"] || "UNKNOWN";
      if (!sku) return;
      demand[mp] = demand[mp] || {};
      demand[mp][sku] = (demand[mp][sku] || 0) + 1;
    });

    /* STOCK */
    const stock = {};
    uni.forEach(r => {
      const sku = r["Sku Code"];
      const binId = r["Shelf"];
      const qty = Number(r["Available (ATP)"]);
      if (!sku || !binId || !qty) return;
      stock[sku] = stock[sku] || {};
      stock[sku][binId] = (stock[sku][binId] || 0) + qty;
    });

    bin.forEach(r => {
      const sku = r["SKU"];
      const binId = r["Bin"];
      const qty = Number(r["Qty"]);
      if (!sku || !binId || !qty) return;
      stock[sku] = stock[sku] || {};
      stock[sku][binId] = (stock[sku][binId] || 0) - qty;
    });

    /* PICK */
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
          mpData[mp].push({ SKU: sku, Style: style, Size: size, BIN: binId, Unit: pick });
          need -= pick;
        }
      }
      mpData[mp].sort(defaultSort);
    }

    buildTabs();
    msg.innerText = "✅ Report generated successfully";

  } catch (err) {
    alert(err);
  }
}

/* ---------- UI ---------- */
function defaultSort(a,b) {
  if(a.SKU!==b.SKU) return a.SKU.localeCompare(b.SKU);
  return SIZE_ORDER.indexOf(a.Size) - SIZE_ORDER.indexOf(b.Size);
}

function buildTabs() {
  mpTabs.innerHTML = "";
  Object.keys(mpData).forEach((mp,i)=>{
    const t = document.createElement("div");
    t.className = "tab" + (i===0?" active":"");
    t.innerText = mp;
    t.onclick = ()=>switchMP(mp);
    mpTabs.appendChild(t);
  });
  switchMP(Object.keys(mpData)[0]);
}

function switchMP(mp) {
  activeMP = mp;
  [...mpTabs.children].forEach(t=>t.classList.toggle("active",t.innerText===mp));
  currentView = [...mpData[mp]];
  applySort();
}

function applySort() {
  const key = sortBy.value;
  if (!key) currentView.sort(defaultSort);
  else if (key==="Size") currentView.sort((a,b)=>SIZE_ORDER.indexOf(a.Size)-SIZE_ORDER.indexOf(b.Size));
  else currentView.sort((a,b)=>a[key].localeCompare(b[key]));
  render();
}

function render() {
  tbody.innerHTML = "";
  currentView.forEach(r=>{
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
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(mpData[mp]),
      mp.substring(0,31)
    );
  }
  XLSX.writeFile(wb, "MP_Wise_Pick_List.xlsx");
}
