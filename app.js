/* =================================================
   Pick List Generator – V3.4
   FULL REPLACEABLE FILE
   NO isCSV / isCsv / isCsvFile ANYWHERE
================================================= */

let mpData = {};
let activeMP = "";
let currentView = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

/* ---------- FILE READER (CSV + EXCEL, NO FLAGS) ---------- */
function readFile(file, sheetName = null) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = e => {
      const workbook = XLSX.read(e.target.result, {
        type: file.name.toLowerCase().endsWith(".csv") ? "string" : "binary"
      });

      const sheet = file.name.toLowerCase().endsWith(".csv")
        ? workbook.Sheets[workbook.SheetNames[0]]
        : sheetName
          ? workbook.Sheets[sheetName]
          : workbook.Sheets[workbook.SheetNames[0]];

      if (!sheet) {
        reject(`Required sheet "${sheetName}" not found in ${file.name}`);
        return;
      }

      resolve(XLSX.utils.sheet_to_json(sheet));
    };

    file.name.toLowerCase().endsWith(".csv")
      ? reader.readAsText(file)
      : reader.readAsBinaryString(file);
  });
}

/* ---------- SKU PARSE ---------- */
function parseSku(sku) {
  if (!sku) return null;
  sku = sku.toString().trim();

  if (!sku.includes("-")) {
    return { style: sku, size: "FS" };
  }

  const parts = sku.split("-");
  return {
    style: parts[0].trim(),
    size: (parts[1] || "FS").trim()
  };
}

/* ---------- UPLOAD STATUS ---------- */
["saleFile", "uniwareFile", "binFile"].forEach(id => {
  document.getElementById(id).addEventListener("change", e => {
    document.getElementById(id.replace("File", "Status")).innerText =
      e.target.files.length ? "✔ Uploaded" : "";
  });
});

/* ---------- GENERATE ---------- */
async function generate() {
  const msg = document.getElementById("msg");
  msg.innerText = "";

  try {
    const saleFile = document.getElementById("saleFile").files[0];
    const uniwareFile = document.getElementById("uniwareFile").files[0];
    const binFile = document.getElementById("binFile").files[0];

    if (!saleFile || !uniwareFile || !binFile) {
      alert("Please upload all 3 files");
      return;
    }

    const sales = await readFile(saleFile);
    const uniware = await readFile(uniwareFile);
    const inward = await readFile(binFile, "Inward");

    /* ---------- DEMAND ---------- */
    const demand = {};
    sales.forEach(r => {
      const sku = r["Item SKU Code (Sku Code)"]?.toString().trim();
      const mp = (r["Channel Name (MP)"] || "UNKNOWN").toString().trim();
      if (!sku) return;

      demand[mp] = demand[mp] || {};
      demand[mp][sku] = (demand[mp][sku] || 0) + 1;
    });

    /* ---------- MASTER STOCK ---------- */
    const stock = {};

    // Uniware → godown
    uniware.forEach(r => {
      const sku = r["Sku Code"]?.toString().trim();
      const qty = Number(r["Available (ATP)"]);
      if (!sku || qty <= 0) return;

      stock[sku] = stock[sku] || {};
      stock[sku]["godown"] = (stock[sku]["godown"] || 0) + qty;
    });

    // Inward → move from godown to bin
    inward.forEach(r => {
      const sku = r["SKU"]?.toString().trim();
      const binId = r["Bin"]?.toString().trim();
      const qty = Number(r["Qty"]);
      if (!sku || !binId || qty <= 0) return;

      stock[sku] = stock[sku] || {};
      stock[sku]["godown"] = (stock[sku]["godown"] || 0) - qty;
      stock[sku][binId] = (stock[sku][binId] || 0) + qty;
    });

    /* ---------- PICK LIST ---------- */
    mpData = {};
    let totalUnits = 0;

    for (const mp in demand) {
      mpData[mp] = [];

      for (const sku in demand[mp]) {
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

        for (const [binId, qty] of bins) {
          if (need <= 0) break;

          const pickQty = Math.min(qty, need);
          mpData[mp].push({
            SKU: sku,
            Style: parsed.style,
            Size: parsed.size,
            BIN: binId,
            Unit: pickQty
          });

          need -= pickQty;
          totalUnits += pickQty;
        }
      }

      mpData[mp].sort((a, b) => {
        if (a.SKU !== b.SKU) return a.SKU.localeCompare(b.SKU);
        return SIZE_ORDER.indexOf(a.Size) - SIZE_ORDER.indexOf(b.Size);
      });
    }

    if (totalUnits === 0) {
      msg.innerText = "⚠️ No pick data generated. Check stock & inward mapping.";
      mpTabs.innerHTML = "";
      tbody.innerHTML = "";
      return;
    }

    buildTabs();
    msg.innerText = `✅ Report generated successfully | Units: ${totalUnits}`;

  } catch (err) {
    alert(err);
  }
}

/* ---------- UI ---------- */
function buildTabs() {
  mpTabs.innerHTML = "";
  const mps = Object.keys(mpData);

  mps.forEach((mp, i) => {
    const tab = document.createElement("div");
    tab.className = "tab" + (i === 0 ? " active" : "");
    tab.innerText = mp;
    tab.onclick = () => switchMP(mp);
    mpTabs.appendChild(tab);
  });

  if (mps.length) switchMP(mps[0]);
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
      </tr>
    `;
  });
}

/* ---------- EXPORT ---------- */
function exportExcel() {
  const wb = XLSX.utils.book_new();
  for (const mp in mpData) {
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(mpData[mp]),
      mp.substring(0, 31)
    );
  }
  XLSX.writeFile(wb, "MP_Wise_Pick_List.xlsx");
}
