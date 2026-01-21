let mpData = {};
let activeMP = "";
let currentView = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

/* ========= SAFE FILE READER (CSV + EXCEL) ========= */
function readFile(file, requiredSheet = null) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = e => {
      const wb = XLSX.read(e.target.result, {
        type: file.name.toLowerCase().endsWith(".csv") ? "string" : "binary"
      });

      let sheet;
      if (file.name.toLowerCase().endsWith(".csv")) {
        sheet = wb.Sheets[wb.SheetNames[0]];
      } else {
        sheet = requiredSheet
          ? wb.Sheets[requiredSheet]
          : wb.Sheets[wb.SheetNames[0]];
      }

      if (!sheet) {
        reject(`Required sheet "${requiredSheet}" not found in ${file.name}`);
        return;
      }

      resolve(XLSX.utils.sheet_to_json(sheet));
    };

    file.name.toLowerCase().endsWith(".csv")
      ? reader.readAsText(file)
      : reader.readAsBinaryString(file);
  });
}

/* ========= SKU PARSE ========= */
function parseSku(sku) {
  sku = sku?.toString().trim();
  if (!sku) return null;
  if (!sku.includes("-")) return { style: sku, size: "FS" };
  const [s, z] = sku.split("-");
  return { style: s.trim(), size: (z || "FS").trim() };
}

/* ========= UPLOAD STATUS ========= */
["saleFile","uniwareFile","binFile"].forEach(id => {
  document.getElementById(id).addEventListener("change", e => {
    document.getElementById(id.replace("File","Status")).innerText =
      e.target.files.length ? "✔ Uploaded" : "";
  });
});

/* ========= GENERATE ========= */
async function generate() {
  const msg = document.getElementById("msg");
  msg.innerText = "";

  try {
    const sales = await readFile(document.getElementById("saleFile").files[0]);
    const uni = await readFile(document.getElementById("uniwareFile").files[0]);
    const bin = await readFile(document.getElementById("binFile").files[0], "Inward");

    /* DEMAND */
    const demand = {};
    sales.forEach(r => {
      const sku = r["Item SKU Code (Sku Code)"]?.toString().trim();
      const mp = (r["Channel Name (MP)"] || "UNKNOWN").toString().trim();
      if (!sku) return;
      demand[mp] = demand[mp] || {};
      demand[mp][sku] = (demand[mp][sku] || 0) + 1;
    });

    /* STOCK */
    const stock = {};

    uni.forEach(r => {
      const sku = r["Sku Code"]?.toString().trim();
      const binId = r["Shelf"]?.toString().trim();
      const qty = Number(r["Available (ATP)"]);
      if (!sku || !binId || qty <= 0) return;
      stock[sku] = stock[sku] || {};
      stock[sku][binId] = (stock[sku][binId] || 0) + qty;
    });

    bin.forEach(r => {
      const sku = r["SKU"]?.toString().trim();
      const binId = r["Bin"]?.toString().trim();
      const qty = Number(r["Qty"]);
      if (!sku || !binId || qty <= 0) return;
      stock[sku] = stock[sku] || {};
      stock[sku][binId] = (stock[sku][binId] || 0) - qty;
    });

    /* PICK */
    mpData = {};
    let total = 0;

    for (let mp in demand) {
      mpData[mp] = [];

      for (let sku in demand[mp]) {
        let need = demand[mp][sku];
        const parsed = parseSku(sku);
        if (!parsed) continue;

        const bins = Object.entries(stock[sku] || {})
          .filter(([_, q]) => q > 0)
          .sort((a,b)=>{
            if (a[0].toLowerCase()==="godown") return -1;
            if (b[0].toLowerCase()==="godown") return 1;
            return a[0].localeCompare(b[0]);
          });

        for (let [binId, qty] of bins) {
          if (need <= 0) break;
          const pick = Math.min(qty, need);
          mpData[mp].push({
            SKU: sku,
            Style: parsed.style,
            Size: parsed.size,
            BIN: binId,
            Unit: pick
          });
          need -= pick;
          total += pick;
        }
      }
    }

    if (total === 0) {
      msg.innerText = "⚠️ No pick data generated. Check SKU match & inward stock.";
      return;
    }

    buildTabs();
    msg.innerText = `✅ Report generated successfully | Units: ${total}`;

  } catch (e) {
    alert(e);
  }
}

/* ========= UI ========= */
function buildTabs() {
  mpTabs.innerHTML = "";
  const mps = Object.keys(mpData);
  if (!mps.length) return;

  mps.forEach((mp,i)=>{
    const t = document.createElement("div");
    t.className = "tab" + (i===0?" active":"");
    t.innerText = mp;
    t.onclick = ()=>switchMP(mp);
    mpTabs.appendChild(t);
  });

  switchMP(mps[0]);
}

function switchMP(mp) {
  activeMP = mp;
  document.querySelectorAll(".tab").forEach(t =>
    t.classList.toggle("active", t.innerText === mp)
  );
  currentView = [...mpData[mp]];
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
