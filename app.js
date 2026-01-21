/* =================================================
   Pick List Generator – FINAL LOGIC VERSION
   FULL REPLACEABLE FILE
================================================= */

let mpData = {};
let shortageData = [];
let activeTab = "";

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

/* ---------- FILE READER ---------- */
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
        reject(`Sheet "${sheetName}" not found in ${file.name}`);
        return;
      }

      resolve(XLSX.utils.sheet_to_json(sheet));
    };

    file.name.toLowerCase().endsWith(".csv")
      ? reader.readAsText(file)
      : reader.readAsBinaryString(file);
  });
}

/* ---------- SKU NORMALIZATION ---------- */
function normalizeSku(rawSku) {
  const original = rawSku?.toString().trim() || "";
  const cleaned = original.replace(/\s+/g, "").toUpperCase();

  let style, size, displaySize;

  if (cleaned.includes("-")) {
    [style, size] = cleaned.split("-");
    displaySize = size;
  } else {
    style = cleaned;
    size = "FS";
    displaySize = "";
  }

  return {
    canonicalSku: `${style}-${size}`,
    style,
    size,
    displaySize
  };
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
  const msg = document.getElementById("msg");
  msg.innerText = "";

  try {
    const sales = await readFile(saleFile.files[0]);
    const uniware = await readFile(uniwareFile.files[0]);
    const inward = await readFile(binFile.files[0], "Inward");

    /* ---------- DEMAND ---------- */
    const demand = {};
    sales.forEach(r => {
      const skuInfo = normalizeSku(r["Item SKU Code (Sku Code)"]);
      const mp = (r["Channel Name (MP)"] || "UNKNOWN").toString().trim();

      demand[mp] = demand[mp] || {};
      demand[mp][skuInfo.canonicalSku] = (demand[mp][skuInfo.canonicalSku] || 0) + 1;
    });

    /* ---------- MASTER STOCK ---------- */
    const stock = {};

    uniware.forEach(r => {
      const skuInfo = normalizeSku(r["Sku Code"]);
      const qty = Number(r["Available (ATP)"]);
      if (qty <= 0) return;

      stock[skuInfo.canonicalSku] = stock[skuInfo.canonicalSku] || {};
      stock[skuInfo.canonicalSku]["godown"] =
        (stock[skuInfo.canonicalSku]["godown"] || 0) + qty;
    });

    inward.forEach(r => {
      const skuInfo = normalizeSku(r["SKU"]);
      const binId = r["Bin"]?.toString().trim();
      const qty = Number(r["Qty"]);
      if (!binId || qty <= 0) return;

      stock[skuInfo.canonicalSku] = stock[skuInfo.canonicalSku] || {};
      stock[skuInfo.canonicalSku]["godown"] =
        Math.max(0, (stock[skuInfo.canonicalSku]["godown"] || 0) - qty);
      stock[skuInfo.canonicalSku][binId] =
        (stock[skuInfo.canonicalSku][binId] || 0) + qty;
    });

    /* ---------- PICK + SHORTAGE ---------- */
    mpData = {};
    shortageData = [];

    for (const mp in demand) {
      mpData[mp] = [];

      for (const sku in demand[mp]) {
        let remaining = demand[mp][sku];
        const bins = Object.entries(stock[sku] || {})
          .filter(([_, q]) => q > 0)
          .sort((a, b) => {
            if (a[0] === "godown") return -1;
            if (b[0] === "godown") return 1;
            return a[0].localeCompare(b[0]);
          });

        const { style, displaySize } = normalizeSku(sku);

        for (const [binId, qty] of bins) {
          if (remaining <= 0) break;
          const pick = Math.min(qty, remaining);

          mpData[mp].push({
            SKU: sku,
            Style: style,
            Size: displaySize,
            BIN: binId,
            Unit: pick
          });

          remaining -= pick;
        }

        if (remaining > 0) {
          shortageData.push({
            SKU: sku,
            MP: mp,
            Units: remaining
          });
        }
      }
    }

    buildTabs();
    msg.innerText = "✅ Report generated successfully";

  } catch (e) {
    alert(e);
  }
}

/* ---------- UI ---------- */
function buildTabs() {
  mpTabs.innerHTML = "";

  Object.keys(mpData).forEach((mp, i) => {
    const tab = document.createElement("div");
    tab.className = "tab" + (i === 0 ? " active" : "");
    tab.innerText = mp;
    tab.onclick = () => switchTab(mp);
    mpTabs.appendChild(tab);
  });

  if (shortageData.length) {
    const tab = document.createElement("div");
    tab.className = "tab";
    tab.innerText = "SHORTAGE";
    tab.onclick = () => switchTab("SHORTAGE");
    mpTabs.appendChild(tab);
  }

  switchTab(Object.keys(mpData)[0]);
}

function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll(".tab").forEach(t =>
    t.classList.toggle("active", t.innerText === tab)
  );

  tbody.innerHTML = "";

  const rows = tab === "SHORTAGE"
    ? shortageData
    : mpData[tab];

  rows.forEach(r => {
    tbody.innerHTML += tab === "SHORTAGE"
      ? `<tr><td>${r.SKU}</td><td></td><td></td><td>${r.MP}</td><td>${r.Units}</td></tr>`
      : `<tr><td>${r.SKU}</td><td>${r.Style}</td><td>${r.Size}</td><td>${r.BIN}</td><td>${r.Unit}</td></tr>`;
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

  if (shortageData.length) {
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(shortageData),
      "SHORTAGE"
    );
  }

  XLSX.writeFile(wb, "MP_Wise_Pick_List.xlsx");
}
