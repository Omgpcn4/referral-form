import { Packer } from "docx";
import { generateReferralFormDocx, buildFilename } from "./docx-generator.js";
import { processAttachmentFile } from "./attachments.js";
import { parseRichText, initRichTextEditors } from "./richtext.js";

const DRAFT_KEY = "referral-form-draft-v2";
const VETS_KEY = "referral-form-vets-v1";

const form = document.getElementById("cert-form");
const dateInput = document.getElementById("date");
const statusMsg = document.getElementById("status-msg");
const clearBtn = document.getElementById("clear-draft-btn");
const generateBtn = document.getElementById("generate-btn");
const generatePdfBtn = document.getElementById("generate-pdf-btn");
const attachmentsInput = document.getElementById("attachments-input");
const attachmentsList = document.getElementById("attachments-list");
const attachmentsStatus = document.getElementById("attachments-status");
const vetSelect = document.getElementById("vetSelect");
const vetNameInput = document.getElementById("vetName");
const licenseNoInput = document.getElementById("licenseNo");
const removeVetBtn = document.getElementById("removeVetBtn");

let attachments = [];

const TEXT_FIELD_IDS = [
  "date",
  "hn",
  "petName",
  "species",
  "gender",
  "breed",
  "age",
  "ownerName",
  "vetName",
  "licenseNo",
];

const RICHTEXT_FIELD_IDS = [
  "history",
  "physicalExam",
  "laboratory",
  "diagnosis",
  "treatment",
  "others",
];

function todayIso() {
  const now = new Date();
  const tzOffsetMs = now.getTimezoneOffset() * 60000;
  return new Date(now - tzOffsetMs).toISOString().slice(0, 10);
}

function collectFormData() {
  const data = {};
  for (const id of TEXT_FIELD_IDS) {
    const el = document.getElementById(id);
    data[id] = el ? el.value : "";
  }
  for (const id of RICHTEXT_FIELD_IDS) {
    const el = document.getElementById(id);
    data[id] = el ? parseRichText(el) : [];
  }
  const purposeEl = form.querySelector('input[name="purpose"]:checked');
  data.purpose = purposeEl ? purposeEl.value : "";
  return data;
}

function collectDraftData() {
  const data = {};
  for (const id of TEXT_FIELD_IDS) {
    const el = document.getElementById(id);
    data[id] = el ? el.value : "";
  }
  for (const id of RICHTEXT_FIELD_IDS) {
    const el = document.getElementById(id);
    data[id] = el ? el.innerHTML : "";
  }
  const purposeEl = form.querySelector('input[name="purpose"]:checked');
  data.purpose = purposeEl ? purposeEl.value : "";
  return data;
}

function applyFormData(data) {
  for (const id of TEXT_FIELD_IDS) {
    const el = document.getElementById(id);
    if (el && data[id] != null) el.value = data[id];
  }
  for (const id of RICHTEXT_FIELD_IDS) {
    const el = document.getElementById(id);
    if (el && data[id] != null) el.innerHTML = data[id];
  }
  if (data.purpose) {
    const purposeEl = form.querySelector(`input[name="purpose"][value="${data.purpose}"]`);
    if (purposeEl) purposeEl.checked = true;
  }
}

function saveDraft() {
  try {
    localStorage.setItem(DRAFT_KEY, JSON.stringify(collectDraftData()));
  } catch (err) {
    // localStorage unavailable (private browsing, quota, etc.) — draft saving is best-effort
  }
}

function loadDraft() {
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function clearDraft() {
  try {
    localStorage.removeItem(DRAFT_KEY);
  } catch (err) {
    // ignore
  }
}

function loadSavedVets() {
  try {
    const raw = localStorage.getItem(VETS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function saveVet(name, licenseNo) {
  const trimmedName = name.trim();
  if (!trimmedName) return;
  try {
    const vets = loadSavedVets();
    const existing = vets.find((v) => v.name === trimmedName);
    if (existing) {
      existing.licenseNo = licenseNo.trim();
    } else {
      vets.push({ name: trimmedName, licenseNo: licenseNo.trim() });
    }
    localStorage.setItem(VETS_KEY, JSON.stringify(vets));
  } catch (err) {
    // localStorage unavailable — saving is best-effort
  }
  populateVetSelect(trimmedName);
}

function removeVet(name) {
  try {
    const vets = loadSavedVets().filter((v) => v.name !== name);
    localStorage.setItem(VETS_KEY, JSON.stringify(vets));
  } catch (err) {
    // localStorage unavailable — removal is best-effort
  }
}

function populateVetSelect(selectedName) {
  const vets = loadSavedVets();
  vetSelect.innerHTML = "";
  const newOption = document.createElement("option");
  newOption.value = "";
  newOption.textContent = "+ New veterinarian…";
  vetSelect.appendChild(newOption);
  for (const vet of vets) {
    const option = document.createElement("option");
    option.value = vet.name;
    option.textContent = vet.licenseNo ? `${vet.name} (${vet.licenseNo})` : vet.name;
    vetSelect.appendChild(option);
  }
  const matched = selectedName && vets.some((v) => v.name === selectedName);
  vetSelect.value = matched ? selectedName : "";
  removeVetBtn.disabled = !matched;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function setStatus(message, type) {
  statusMsg.textContent = message;
  statusMsg.classList.remove("error", "success");
  if (type) statusMsg.classList.add(type);
}

function renderAttachmentsList() {
  attachmentsList.innerHTML = "";
  for (const att of attachments) {
    const li = document.createElement("li");
    li.className = "attachment-item";

    const img = document.createElement("img");
    img.src = att.thumbnail;
    img.alt = att.label;
    li.appendChild(img);

    const name = document.createElement("span");
    name.className = "attachment-name";
    name.textContent = att.label;
    name.title = att.label;
    li.appendChild(name);

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "attachment-remove";
    removeBtn.textContent = "×";
    removeBtn.setAttribute("aria-label", `Remove ${att.label}`);
    removeBtn.addEventListener("click", () => {
      attachments = attachments.filter((a) => a.id !== att.id);
      renderAttachmentsList();
    });
    li.appendChild(removeBtn);

    attachmentsList.appendChild(li);
  }
}

async function handleAttachmentFiles(fileList) {
  const files = Array.from(fileList);
  if (files.length === 0) return;
  attachmentsStatus.textContent = `Processing ${files.length} file(s)…`;
  for (const file of files) {
    try {
      const processed = await processAttachmentFile(file);
      attachments.push(...processed);
    } catch (err) {
      console.error(err);
      attachmentsStatus.textContent = `Failed to process ${file.name}: ${err.message}`;
    }
  }
  renderAttachmentsList();
  if (!attachmentsStatus.textContent.startsWith("Failed")) {
    attachmentsStatus.textContent = attachments.length
      ? `${attachments.length} attachment page(s) ready.`
      : "";
  }
}

function init() {
  initRichTextEditors();

  const draft = loadDraft();
  if (draft) {
    applyFormData(draft);
  }
  if (!dateInput.value) {
    dateInput.value = todayIso();
  }

  populateVetSelect(vetNameInput.value);

  vetSelect.addEventListener("change", () => {
    const vets = loadSavedVets();
    const selected = vets.find((v) => v.name === vetSelect.value);
    removeVetBtn.disabled = !selected;
    if (selected) {
      vetNameInput.value = selected.name;
      licenseNoInput.value = selected.licenseNo;
      saveDraft();
    }
  });

  removeVetBtn.addEventListener("click", () => {
    const name = vetSelect.value;
    if (!name) return;
    if (!confirm(`Remove "${name}" from the saved veterinarian list?`)) return;
    removeVet(name);
    populateVetSelect("");
  });

  form.addEventListener("input", saveDraft);
  form.addEventListener("change", saveDraft);

  attachmentsInput.addEventListener("change", () => {
    handleAttachmentFiles(attachmentsInput.files);
    attachmentsInput.value = "";
  });

  clearBtn.addEventListener("click", () => {
    if (!confirm("Clear all fields and delete the saved draft?")) return;
    form.reset();
    for (const id of RICHTEXT_FIELD_IDS) {
      document.getElementById(id).innerHTML = "";
    }
    clearDraft();
    dateInput.value = todayIso();
    attachments = [];
    renderAttachmentsList();
    attachmentsStatus.textContent = "";
    setStatus("Form cleared.", "success");
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    generateBtn.disabled = true;
    setStatus("Generating document…");
    try {
      const data = collectFormData();
      data.attachments = attachments;
      saveVet(data.vetName, data.licenseNo);
      const doc = await generateReferralFormDocx(data);
      const blob = await Packer.toBlob(doc);
      const filename = buildFilename(data);
      downloadBlob(blob, filename);
      setStatus(`Downloaded ${filename}`, "success");
    } catch (err) {
      console.error(err);
      setStatus("Failed to generate document. See console for details.", "error");
    } finally {
      generateBtn.disabled = false;
    }
  });

  generatePdfBtn.addEventListener("click", async () => {
    generatePdfBtn.disabled = true;
    setStatus("Generating PDF…");
    try {
      const { generateReferralFormPdf } = await import("./pdf-generator.js");
      const data = collectFormData();
      data.attachments = attachments;
      saveVet(data.vetName, data.licenseNo);
      const pdfBytes = await generateReferralFormPdf(data);
      const blob = new Blob([pdfBytes], { type: "application/pdf" });
      const filename = buildFilename(data).replace(/\.docx$/, ".pdf");
      downloadBlob(blob, filename);
      setStatus(`Downloaded ${filename}`, "success");
    } catch (err) {
      console.error(err);
      setStatus("Failed to generate PDF. See console for details.", "error");
    } finally {
      generatePdfBtn.disabled = false;
    }
  });
}

init();
