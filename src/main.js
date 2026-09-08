import { Packer } from "docx";
import { generateHealthCertificateDocx, buildFilename } from "./docx-generator.js";

const DRAFT_KEY = "referral-form-draft-v1";

const form = document.getElementById("cert-form");
const dateInput = document.getElementById("date");
const statusMsg = document.getElementById("status-msg");
const clearBtn = document.getElementById("clear-draft-btn");
const generateBtn = document.getElementById("generate-btn");

const FIELD_IDS = [
  "recipientName",
  "date",
  "ownerName",
  "ownerAddress",
  "ownerPhone",
  "animalName",
  "hn",
  "species",
  "gender",
  "ageYears",
  "ageMonths",
  "ageDays",
  "breed",
  "microchip",
  "history",
  "physicalExam",
  "laboratory",
  "diagnosis",
  "treatment",
  "others",
  "vetName",
  "licenseNo",
];

function todayIso() {
  const now = new Date();
  const tzOffsetMs = now.getTimezoneOffset() * 60000;
  return new Date(now - tzOffsetMs).toISOString().slice(0, 10);
}

function collectFormData() {
  const data = {};
  for (const id of FIELD_IDS) {
    const el = document.getElementById(id);
    data[id] = el ? el.value : "";
  }
  return data;
}

function applyFormData(data) {
  for (const id of FIELD_IDS) {
    const el = document.getElementById(id);
    if (el && data[id] != null) el.value = data[id];
  }
}

function saveDraft() {
  try {
    localStorage.setItem(DRAFT_KEY, JSON.stringify(collectFormData()));
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

function setStatus(message, type) {
  statusMsg.textContent = message;
  statusMsg.classList.remove("error", "success");
  if (type) statusMsg.classList.add(type);
}

function init() {
  const draft = loadDraft();
  if (draft) {
    applyFormData(draft);
  }
  if (!dateInput.value) {
    dateInput.value = todayIso();
  }

  form.addEventListener("input", saveDraft);

  clearBtn.addEventListener("click", () => {
    if (!confirm("Clear all fields and delete the saved draft?")) return;
    form.reset();
    clearDraft();
    dateInput.value = todayIso();
    setStatus("Form cleared.", "success");
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    generateBtn.disabled = true;
    setStatus("Generating document…");
    try {
      const data = collectFormData();
      const doc = generateHealthCertificateDocx(data);
      const blob = await Packer.toBlob(doc);
      const filename = buildFilename(data);

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setStatus(`Downloaded ${filename}`, "success");
    } catch (err) {
      console.error(err);
      setStatus("Failed to generate document. See console for details.", "error");
    } finally {
      generateBtn.disabled = false;
    }
  });
}

init();
