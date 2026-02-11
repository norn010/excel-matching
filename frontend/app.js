const API_BASE = (typeof window !== "undefined" && window.location.port === "8000") ? "" : "http://localhost:8000";

const esgInput = document.getElementById("esg-file");
const taxInput = document.getElementById("tax-file");
const esgName = document.getElementById("esg-name");
const taxName = document.getElementById("tax-name");
const sheetNameInput = document.getElementById("sheet-name");
const btnMatch = document.getElementById("btn-match");
const apiStatus = document.getElementById("api-status");
const resultSection = document.getElementById("result-section");
const summaryEl = document.getElementById("summary");
const resultThead = document.getElementById("result-thead");
const resultBody = document.getElementById("result-body");

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s ?? "";
  return d.innerHTML;
}

function setFileName(el, file) {
  el.textContent = file ? file.name : "ยังไม่ได้เลือก";
}

function checkCanMatch() {
  btnMatch.disabled = !(esgInput.files?.length && taxInput.files?.length);
}

esgInput.addEventListener("change", () => {
  setFileName(esgName, esgInput.files?.[0]);
  checkCanMatch();
});
taxInput.addEventListener("change", () => {
  setFileName(taxName, taxInput.files?.[0]);
  checkCanMatch();
});

async function callMatch() {
  const esgFile = esgInput.files?.[0];
  const taxFile = taxInput.files?.[0];
  if (!esgFile || !taxFile) return;

  btnMatch.disabled = true;
  apiStatus.className = "status loading";
  apiStatus.textContent = "กำลังเทียบข้อมูล...";

  const form = new FormData();
  form.append("esg_file", esgFile);
  form.append("tax_file", taxFile);

  const sheetName = (sheetNameInput.value || "ตารางไมวัน").trim();
  const esgSheetName = "เงินรางวัลเรียกเก็บ ESG";
  const matchKeyCol = 3; // เลขถัง
  const esgCols = "2,3,5,6,8,14,15"; // ชีตเงินรางวัลเรียกเก็บ ESG
  const taxCols = "1,2,4,5,7,13,14"; // ชีตตารางไมวัน
  const url = `${API_BASE}/match-columns?sheet_name=${encodeURIComponent(sheetName)}&esg_sheet_name=${encodeURIComponent(esgSheetName)}&esg_cols=${encodeURIComponent(esgCols)}&tax_cols=${encodeURIComponent(taxCols)}&match_key_col=${matchKeyCol}&case_sensitive=true`;

  try {
    const res = await fetch(url, { method: "POST", body: form });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      apiStatus.className = "status error";
      apiStatus.textContent = data.detail || "เกิดข้อผิดพลาด";
      return;
    }

    apiStatus.className = "status success";
    apiStatus.textContent = "เทียบข้อมูลเสร็จแล้ว";

    const { column_labels, match_key_label, summary, results } = data;
    summaryEl.innerHTML = `
      <p>รวมแถว: <strong>${summary.total_rows}</strong> · ตรงทุกคอลัมน์ (แถวเขียว): <strong class="ok">${summary.all_match}</strong> · มีคอลัมน์ไม่ตรง (แถวเหลือง): <strong class="warn">${summary.has_mismatch}</strong></p>
      <p class="hint">เงื่อนไขจับคู่แถว: <strong>${escapeHtml(match_key_label || "เลขถัง")}</strong> เป็นตัวหลัก</p>
    `;

    resultThead.innerHTML = "<th>ลำดับ</th>" + (column_labels || []).map((x) => `<th>${escapeHtml(x)}</th>`).join("") + "<th>ข้อมูลที่ไม่ตรง</th>";
    resultBody.innerHTML = (results || []).map((r) => {
      const rowClass = r.all_match ? "row-ok" : "row-warn";
      const cells = (r.values_esg || []).map((v, i) => {
        const ok = r.cells_match && r.cells_match[i];
        const tv = r.values_tax && r.values_tax[i] !== undefined ? r.values_tax[i] : "";
        const title = !ok ? `ESG: ${v} | ภาษี: ${tv}` : "";
        return `<td class="${ok ? "cell-ok" : "cell-warn"}"${title ? ` title="${escapeHtml(title)}"` : ""}>${escapeHtml(v)}</td>`;
      }).join("");
      const mismatchHtml = (r.mismatches || []).length
        ? r.mismatches.map((m) => `<span class="mismatch-label">${escapeHtml(m.column_label)}</span>: ESG=<strong>${escapeHtml(m.esg_value || "(ว่าง)")}</strong> ภาษี=<strong>${escapeHtml(m.tax_value || "(ว่าง)")}</strong>`).join(" · ")
        : "—";
      return `<tr class="${rowClass}"><td>${r.row}</td>${cells}<td class="mismatch-cell">${mismatchHtml}</td></tr>`;
    }).join("");
    resultSection.hidden = false;
  } catch (e) {
    apiStatus.className = "status error";
    apiStatus.textContent = "เชื่อมต่อ API ไม่ได้";
  } finally {
    btnMatch.disabled = false;
  }
}

btnMatch.addEventListener("click", callMatch);
