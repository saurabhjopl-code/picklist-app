let pickList = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

function readExcel(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const workbook = XLSX.read(e.target.result, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      resolve(XLSX.utils.sheet_to_json(sheet));
    };
    reader.readAsBinaryString(file);
  });
}

function parseSku(sku) {
  if (!sku.includes("-")) {
    return { style: sku, size: "FS" };
  }
  const parts = sku.split("-");
  return { style: parts[0], size: parts[1] || "FS" };
}

async function generatePickList() {
  const orderFile = document.getElementById("orderFile").files[0];
  const stockFile = document.getElementById("stockFile").files[0];
  if (!orderFile || !stockFile) return alert("Upload both files");

  const orders = await readExcel(orderFile);
  const stock = await readExcel(stockFile);

  pickList = [];

  orders.forEach(order => {
    const sku = order["Sku Code"];
    let qtyNeeded = Number(order["Order Units"]);
    if (!sku || qtyNeeded <= 0) return;

    const { style, size } = parseSku(sku);

    let stockRows = stock
      .filter(s => s["Sku Code"] === sku && Number(s["Available (ATP)"]) > 0)
      .sort((a, b) => {
        if (a["Shelf"].toLowerCase() === "godown") return -1;
        if (b["Shelf"].toLowerCase() === "godown") return 1;
        return a["Shelf"].localeCompare(b["Shelf"]);
      });

    for (let s of stockRows) {
      if (qtyNeeded <= 0) break;

      const available = Number(s["Available (ATP)"]);
      const pickQty = Math.min(available, qtyNeeded);

      pickList.push({
        SKU: sku,
        Style: style,
        Size: size,
        BIN: s["Shelf"],
        Unit: pickQty
      });

      qtyNeeded -= pickQty;
    }
  });

  pickList.sort((a, b) => {
    if (a.SKU !== b.SKU) return a.SKU.localeCompare(b.SKU);
    return SIZE_ORDER.indexOf(a.Size) - SIZE_ORDER.indexOf(b.Size);
  });

  renderTable();
}

function renderTable() {
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  pickList.forEach(r => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.SKU}</td>
      <td>${r.Style}</td>
      <td>${r.Size}</td>
      <td>${r.BIN}</td>
      <td>${r.Unit}</td>
    `;
    tbody.appendChild(tr);
  });
}

function exportExcel() {
  if (!pickList.length) return alert("No data");

  const ws = XLSX.utils.json_to_sheet(
    pickList.map(r => ({
      SKU: r.SKU,
      "Style ID": r.Style,
      Size: r.Size,
      BIN: r.BIN,
      Unit: r.Unit
    }))
  );

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "PickList");
  XLSX.writeFile(wb, "Pick_List.xlsx");
}
