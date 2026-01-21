let masterPick = [];
let filteredPick = [];

const SIZE_ORDER = [
  "FS","XS","S","M","L","XL","XXL",
  "3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"
];

function readXlsx(file) {
  return new Promise(res => {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const sh = wb.Sheets[wb.SheetNames[0]];
      res(XLSX.utils.sheet_to_json(sh));
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
  const bin = await readXlsx(binFile.files[0]);

  const demand = {};
  sale.forEach(r => {
    const sku = r["Item SKU Code (Sku Code)"];
    if (sku) demand[sku] = (demand[sku] || 0) + 1;
  });

  const stock = {};

  function addStock(arr, sign) {
    arr.forEach(r => {
      const sku = r["Sku Code"];
      const binId = r["Shelf"];
      const qty = Number(r["Available (ATP)"]) * sign;
      if (!sku || !binId || !qty) return;

      stock[sku] = stock[sku] || {};
      stock[sku][binId] = (stock[sku][binId] || 0) + qty;
    });
  }

  addStock(uni, 1);
  addStock(bin, -1);

  masterPick = [];

  for (let sku in demand) {
    let need = demand[sku];
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

      masterPick.push({
        SKU: sku,
        Style: style,
        Size: size,
        BIN: binId,
        Unit: pick
      });

      need -= pick;
    }
  }

  masterPick.sort((a,b)=>{
    if(a.SKU!==b.SKU) return a.SKU.localeCompare(b.SKU);
    return SIZE_ORDER.indexOf(a.Size) - SIZE_ORDER.indexOf(b.Size);
  });

  filteredPick = [...masterPick];
  render();
}

function applyFilters() {
  const s = fSku.value.toLowerCase();
  const st = fStyle.value.toLowerCase();
  const sz = fSize.value.toLowerCase();
  const b = fBin.value.toLowerCase();

  filteredPick = masterPick.filter(r =>
    r.SKU.toLowerCase().includes(s) &&
    r.Style.toLowerCase().includes(st) &&
    r.Size.toLowerCase().includes(sz) &&
    r.BIN.toLowerCase().includes(b)
  );
  render();
}

function render() {
  tbody.innerHTML = "";
  filteredPick.forEach(r => {
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
  if (!filteredPick.length) return alert("No data");
  const ws = XLSX.utils.json_to_sheet(filteredPick);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "PickList");
  XLSX.writeFile(wb, "Pick_List.xlsx");
}
