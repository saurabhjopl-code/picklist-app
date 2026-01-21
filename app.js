let mpData = {};
let activeMP = "";
let currentView = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

/* ========= SAFE FILE READER ========= */
function readFile(file, requiredSheet = null) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const isCSV = file.name.toLowerCase().endsWith(".csv");
      const wb = XLSX.read(e.target.result, { type: isCSV ? "string" : "binary" });

      let sheet;
      if (isCSV) {
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

    isCSV ? reader.readAsText(file) : reader.readAsBinaryString(file);
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

/* ========= GENERATE ========= */
async function generate() {
  const msg = document.getElementById("msg");
  msg.innerText = "";

  try {
    const sales = await readFile(document.getElementById("saleFile").files[0]);
    const uni = await readFile(document.getElementById("uniwareFile").files[0]);
    const inward = await readFile(document.getElementById("binFile").files[0], "Inward");

    /* -------- DEMAND (MP → SKU) -------- */
    const demand = {};
    sales.forEach(r => {
      const sku = r["Item SKU Code (Sku Code)"]?.toString().trim();
      const mp = (r["Channel Name (MP)"] || "UNKNOWN").toString().trim();
      if (!sku) return;

      demand[mp] = demand[mp] || {};
      demand[mp][sku] = (demand[mp][sku] || 0) + 1;
    });

    /* -------- BUILD MASTER STOCK -------- */
    const stock = {};

    // Step 1: Load Uniware stock into godown
    uni.forEach(r => {
      const sku = r["Sku Code"]?.toString().trim();
      const qty = Number(r["Available (ATP)"]);
      if (!sku || qty <= 0) return;

      stock[sku] = stock[sku] || {};
      stock[sku]["godown"] = (stock[sku]["godown"] || 0) + qty;
    });

    // Step 2: Move inward from godown → bins
    inward.forEach(r => {
      const sku = r["SKU"]?.toString().trim();
      const binId = r["Bin"]?.toString().trim();
      const qty = Number(r["Qty"]);
      if (!sku || !binId || qty <= 0) return;

      stock[sku] = stock[sku] || {};
      stock[sku]["godown"] = (stock[sku]["godown"] || 0) - qty;
      stock[sku][binId] = (stock[sku][binId] || 0) + qty;
    });

    /* -------- PICK LIST -------- */
    mpData = {};
    let totalUnits = 0;

    for (let mp in demand) {
      mpData[mp] = [];

      for (let sku in demand[mp]) {
        let need = demand[mp][sku];
        const parsed = parseSku(sku);
        if (!parsed) continue;

        const bins = Object.entries(stock[sku] || {})
          .filter(([_, q]) => q > 0)
          .sort((a, b) => {
            if (a[0] === "godown") return -1;
            if (b[0] === "godown") return 1;
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
          totalUnits += pick;
        }
      }

      mpData[mp].sort((a, b) => {
        if (a.SKU !== b.SKU) return a.SKU.localeCompare(b.SKU);
        return SIZE_ORDER.indexOf(a.Size) - SIZE_ORDER.indexOf(b.Size);
      });
    }

    if (totalUnits === 0) {
      msg.innerText = "⚠️ No pick data generated even after master stock logic.";
      return;
    }

    buildTabs();
    msg.innerText = `✅ Report generated successfully | Units: ${totalUnits}`;

  } catch (e) {
    alert(e);
  }
}

/* ========= UI ========= */
function buildTabs() {
  mpTabs.innerHTML = "";
  const mps = Object.keys(mpData);
  if (!mps.length) return;

  mps.forEach((mp, i) => {
    const t = document.createElement("div");
    t.className = "tab" + (i === 0 ? " active" : "");
    t.innerText = mp;
    t.onclick = () => switchMP(mp);
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
  currentView.forEach(r => {
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
      mp.substring(0, 31)
    );
  }
  XLSX.writeFile(wb, "MP_Wise_Pick_List.xlsx");
}
