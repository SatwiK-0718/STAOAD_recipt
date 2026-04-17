// let dataRows = [];

// ✅ Convert S.No to PURE NUMBER
function getSerialNumber(value) {
  if (!value) return "";

  // If already number
  if (typeof value === "number") {
    return Math.floor(value);
  }

  // If string like "Mar-26"
  if (typeof value === "string") {
    const num = value.match(/\d+/);
    return num ? parseInt(num[0]) : "";
  }

  // If Excel date object
  const date = new Date(value);
  if (!isNaN(date)) {
    return date.getDate();
  }

  return "";
}

// Read Excel
document.getElementById("xlInput").addEventListener("change", function (e) {
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const wb = XLSX.read(data, { type: "array" });

    const sheet = wb.Sheets[wb.SheetNames[0]];

    const raw = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true,
    });

    // ✅ MAP S.No → Receipt No (numeric)
    dataRows = raw
      .filter((r) => r["Owner"] && r["Flat No."])
      .map((r) => ({
        ...r,
        "Receipt No.": getSerialNumber(r["S.No"]),
      }));

    if (dataRows.length) {
      document.getElementById("dlBtn").disabled = false;
      buildHTML();
    } else {
      alert("No valid data found");
    }
  };

  reader.readAsArrayBuffer(e.target.files[0]);
});

// Format date → dd/mm/yy
function formatDate(value) {
  if (!value) return "";

  let date;

  if (value instanceof Date) {
    date = value;
  } else if (typeof value === "number") {
    const excelEpoch = new Date(1899, 11, 30);
    date = new Date(excelEpoch.getTime() + value * 86400000);
  } else {
    date = new Date(value);
  }

  if (isNaN(date)) return "";

  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);

  return `${day}/${month}/${year}`;
}

// Build Receipts (3 per page)
function buildHTML() {
  const zone = document.getElementById("render-zone");
  zone.innerHTML = "";

  const today = new Date().toLocaleDateString("en-GB");

  let pageDiv;

  dataRows.forEach((row, index) => {
    if (index % 3 === 0) {
      pageDiv = document.createElement("div");
      pageDiv.className = "page";
      zone.appendChild(pageDiv);
    }

    const receipt = document.createElement("div");
    receipt.className = "receipt";

    receipt.innerHTML = `
      <div class="header">
          <h2>SRI TRIGUNATMIKA APARTMENT OWNERS ASSOCIATION</h2>
          <p>Saptapur, Haliyal Road, Bharati Nagar, Dharwad (Reg. No. DWR-S5-2025-16)</p>
      </div>
      <br><br>

      <div class="top-meta">
          <span class="receipt-title">RECEIPT</span>
          <span class="date-line">Date: <span class="val">${today}</span></span>
      </div>

      <div class="body-text">
          <div class="row-grid">
              <span>Receipt No: <span class="val">${row["Receipt No."]}</span></span>
              <span>Flat No: <span class="val">${row["Flat No."]}</span></span>
          </div>
          <br><br> 

          <div>Received with thanks towards Sinking/Maintenance/Registration charges from</div>           
          
          <div>
              Owner's Name:
              <span class="val">${row["Owner"]}</span>
          </div>

          <br><br>

          <div class="row-grid">
              <span>Amount: <span class="val">₹ ${row["Amount"]}</span></span>
              <span>Mode: <span class="val">${row["Mode of Pyt"]}</span></span>
          </div>

          <div class="row-grid">
              <span>Ref: <span class="val">${row["Chq No/Trf date"]}</span></span>
              <span>Dated: <span class="val">${formatDate(row["Trfr/Chq date"])}</span></span>
          </div>

          <div>Period: <span class="val">${row["Period"]}</span></div>
      </div>

      <br>

      <div class="footer">
          <div class="sig-box">
              For STAOAD<br>
              Secretary / Treasurer
          </div>
      </div>
    `;

    pageDiv.appendChild(receipt);
  });
}

// Download PDF
document.getElementById("dlBtn").addEventListener("click", function () {
  const element = document.getElementById("render-zone");
  element.style.display = "block";

  html2pdf()
    .set({
      margin: [0, 0, 0, 0],
      filename: "STAOAD_Receipts.pdf",
      html2canvas: { scale: 2 },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["avoid-all", "css", "legacy"] },
    })
    .from(element)
    .save()
    .then(() => {
      element.style.display = "none";
    });
});
