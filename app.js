/* =================================================
   Pick List Generator – FINAL STABLE VERSION
================================================= */

let mpData = {};
let shortageData = [];
let activeTab = "";
let currentView = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

/* ---------- MP MAPPING ---------- */
const MP_MAP = {
  "MYNTRAPPMP": "MYNTRA",
  "MYNTRAPPMP-VIHAAN": "MYNTRA",
  "NYKAA_FASHION_NEW": "NYKAA",
  "SNAP_DEAL": "SNAPDEAL",
  "SNAPDEAL_WW": "SNAPDEAL",
  "SNAPDEAL_ARF": "SNAPDEAL",
  "SNAPDEAL_SGA": "SNAPDEAL",
  "SNAPDEAL_SVF": "SNAPDEAL",
  "SNAPDEAL_VEXIM": "SNAPDEAL",
  "MEESHO": "MEESHO",
  "MEESHO_WW": "MEESHO",
  "MEESHO_SVF": "MEESHO",
  "MEESHO_VEXIM": "MEESHO",
  "MEESHO_ARF": "MEESHO",
  "MEESHO_SGA": "MEESHO",
  "FLIPKART": "FLIPKART",
  "FLIPKART_SVF": "FLIPKART",
  "FLIPKART_ARF": "FLIPKART",
  "FLIPKART_WW": "FLIPKART",
  "FLIPKART_VEXIM": "FLIPKART",
  "FLIPKART_SGA": "FLIPKART",
  "AMAZON_FLEX_API": "AMAZON",
  "AMAZON_IN_API": "AMAZON",
  "AJIO_NEW": "AJIO",
  "TATACLIQ": "TATACLIQ",
  "LIMEROAD": "LIMEROAD",
  "MIRRAW": "MIRRAW",
  "RJN_SHOPIFY": "SHOPSY"
};

/* ---------- FILE READER ---------- */
function readFile(file, sheetName = null) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, {
        type: file.name.toLowerCase().endsWith(".csv") ? "string" : "binary"
      });

      const sheet = file.name.toLowerCase().endsWith(".csv")
        ? wb.Sheets[wb.SheetNames[0]]
        : sheetName
          ? wb.Sheets[sheetName]
          : wb.Sheets[wb.SheetNames[0]];

      if (!sheet) {
        reject(`Sheet "${sheetName}" not found`);
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
  if (!rawSku || !rawSku.toString().trim()) return null;

  const cleaned = rawSku.toString().trim().toUpperCase().replace(/\s+/g, "");
  let style, size, displaySize;

  if (cleaned.includes("-")) {
    [style, size] = cleaned.split("-");
    displaySize = size;
  } else {
    style = cleaned;
    size = "FS";
    displaySize = "";
  }

  if (!style) return null;

  return {
    canonicalSku: `${style}-${size}`,
    style,
    size,
    displaySize
  };
}

/* ---------- SORT ---------- */
function applySort() {
  const key = document.getElementById("sortBy").value;

  if (!key) {
    currentView.sort(defaultSort);
  } else if (key === "Size") {
    currentView.sort((a, b) =>
      SIZE_ORDER.indexOf(a.Size || "FS") - SIZE_ORDER.indexOf(b.Size || "FS")
    );
  } else {
    currentView.sort((a, b) =>
      (a[key] || "").localeCompare(b[key] || "")
    );
  }

  render();
}

function defaultSort(a, b) {
  if (a.SKU !== b.SKU) return a.SKU.localeCompare(b.SKU);
  return SIZE_ORDER.indexOf(a.Size || "FS") - SIZE_ORDER.indexOf(b.Size || "FS");
}

/* ---------- GENERATE ---------- */
async function generate() {
  const msg = document.getElementById("msg");
  msg.innerText = "";

  mpData = {};
  shortageData = [];

  try {
    const sales = await readFile(saleFile.files[0]);
    const uniware = await readFile(uniwareFile.files[0]);
    const inward = await readFile(binFile.files[0], "Inward");

    /* ---------- DEMAND ---------- */
    const demand = {};

    sales.forEach(r => {
      const skuInfo = normalizeSku(r["Item SKU Code (Sku Code)"]);
      const rawMp = r["Channel Name"] || r["Channel Name (MP)"] || "";
      const cleanedMp = rawMp.toString().trim().toUpperCase().replace(/\s+/g, "_");
      const mp = MP_MAP[cleanedMp] || "UNKNOWN";

      if (!skuInfo) {
        shortageData.push({ SKU: "", MP: mp, Units: 1 });
        return;
      }

      demand[mp] = demand[mp] || {};
      demand[mp][skuInfo.canonicalSku] =
        (demand[mp][skuInfo.canonicalSku] || 0) + 1;
    });

    /* ---------- MASTER STOCK ---------- */
    const stock = {};

    uniware.forEach(r => {
      const skuInfo = normalizeSku(r["Sku Code"]);
      const qty = Number(r["Available (ATP)"]);
      if (!skuInfo || qty <= 0) return;

      stock[skuInfo.canonicalSku] = stock[skuInfo.canonicalSku] || {};
      stock[skuInfo.canonicalSku]["godown"] =
        (stock[skuInfo.canonicalSku]["godown"] || 0) + qty;
    });

    inward.forEach(r => {
      const skuInfo = normalizeSku(r["SKU"]);
      const binId = r["Bin"]?.toString().trim();
      const qty = Number(r["Qty"]);
      if (!skuInfo || !binId || qty <= 0) return;

      stock[skuInfo.canonicalSku] = stock[skuInfo.canonicalSku] || {};
      stock[skuInfo.canonicalSku]["godown"] =
        Math.max(0, (stock[skuInfo.canonicalSku]["godown"] || 0) - qty);
      stock[skuInfo.canonicalSku][binId] =
        (stock[skuInfo.canonicalSku][binId] || 0) + qty;
    });

    /* ---------- PICK + SHORTAGE ---------- */
    for (const mp in demand) {
      mpData[mp] = [];

      for (const sku in demand[mp]) {
        let need = demand[mp][sku];
        const bins = Object.entries(stock[sku] || {})
          .filter(([_, q]) => q > 0)
          .sort((a, b) => {
            if (a[0] === "godown") return -1;
            if (b[0] === "godown") return 1;
            return a[0].localeCompare(b[0]);
          });

        const { style, displaySize } = normalizeSku(sku);

        for (const [binId, qty] of bins) {
          if (need <= 0) break;
          const pick = Math.min(qty, need);

          mpData[mp].push({
            SKU: sku,
            Style: style,
            Size: displaySize,
            BIN: binId,
            Unit: pick
          });

          need -= pick;
        }

        if (need > 0) {
          shortageData.push({ SKU: sku, MP: mp, Units: need });
        }
      }

      mpData[mp].sort(defaultSort);
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

  currentView = tab === "SHORTAGE" ? shortageData : mpData[tab];
  applySort();
}

function render() {
  tbody.innerHTML = "";

  currentView.forEach(r => {
    tbody.innerHTML += activeTab === "SHORTAGE"
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
