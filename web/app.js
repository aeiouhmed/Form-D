const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const fileInfo = document.getElementById("file-info");
const convertForm = document.getElementById("convert-form");
const convertButton = document.getElementById("convert-button");
const statusEl = document.getElementById("status");
const lastFileSection = document.getElementById("last-file");
const lastFileLink = document.getElementById("last-file-link");
const uomSelect = document.getElementById("uom-select");

let selectedFile = null;
let lastDownloadUrl = null;

const ACCEPTED_TYPES = [
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
];
const ACCEPTED_EXTENSIONS = [".xls", ".xlsx"];
const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20 MB

const spinner = document.createElement("div");
spinner.className = "spinner hidden";
spinner.setAttribute("role", "status");
spinner.setAttribute("aria-live", "polite");
spinner.innerHTML =
  '<span class="spinner-icon" aria-hidden="true"></span><span class="spinner-text">Converting…</span>';
convertForm.insertAdjacentElement("afterend", spinner);

const showSpinner = () => spinner.classList.remove("hidden");
const hideSpinner = () => spinner.classList.add("hidden");

const resetStatus = () => {
  statusEl.textContent = "";
  statusEl.classList.remove("error", "success");
};

const setStatus = (message, type = "info") => {
  statusEl.textContent = message;
  statusEl.classList.remove("error", "success");
  if (type === "error") {
    statusEl.classList.add("error");
  }
  if (type === "success") {
    statusEl.classList.add("success");
  }
};

const formatSize = (bytes) => {
  if (!Number.isFinite(bytes)) return "";
  const thresholds = [
    { unit: "MB", value: 1024 ** 2 },
    { unit: "KB", value: 1024 },
  ];
  for (const { unit, value } of thresholds) {
    if (bytes >= value) {
      return `${(bytes / value).toFixed(1)} ${unit}`;
    }
  }
  return `${bytes} bytes`;
};

const isValidFile = (file) => {
  if (!file) return false;
  const hasValidType = ACCEPTED_TYPES.includes(file.type);
  const hasValidExtension = ACCEPTED_EXTENSIONS.some((ext) =>
    file.name.toLowerCase().endsWith(ext)
  );
  return hasValidType || hasValidExtension;
};

const updateFileInfo = () => {
  if (!selectedFile) {
    fileInfo.textContent = "No file selected yet.";
    convertButton.disabled = true;
    return;
  }
  fileInfo.textContent = `${selectedFile.name} (${formatSize(
    selectedFile.size
  )})`;
  convertButton.disabled = false;
};

const resetFileSelection = (message) => {
  selectedFile = null;
  fileInput.value = "";
  updateFileInfo();
  if (message) {
    setStatus(message, "error");
  }
};

const handleFiles = (files) => {
  resetStatus();
  const file = files?.[0];
  if (!file) {
    return;
  }
  if (!isValidFile(file)) {
    resetFileSelection(
      "Please select a valid Excel file (.xls or .xlsx)."
    );
    return;
  }
  if (file.size > MAX_FILE_SIZE) {
    resetFileSelection(
      `The selected file is ${formatSize(
        file.size
      )}. Files larger than 20 MB are not supported.`
    );
    return;
  }
  selectedFile = file;
  updateFileInfo();
};

const onDragEnter = (event) => {
  event.preventDefault();
  event.stopPropagation();
  dropzone.classList.add("drag-active");
};

const onDragOver = (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = "copy";
  }
};

const onDragLeave = (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!dropzone.contains(event.relatedTarget)) {
    dropzone.classList.remove("drag-active");
  }
};

const onDrop = (event) => {
  event.preventDefault();
  event.stopPropagation();
  dropzone.classList.remove("drag-active");
  const { files } = event.dataTransfer || {};
  handleFiles(files);
};

dropzone.addEventListener("dragenter", onDragEnter);
dropzone.addEventListener("dragover", onDragOver);
dropzone.addEventListener("dragleave", onDragLeave);
dropzone.addEventListener("drop", onDrop);

fileInput.addEventListener("change", (event) => {
  handleFiles(event.target.files);
});

dropzone.addEventListener("click", () => fileInput.click());

convertForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!selectedFile) {
    setStatus("Select a file before converting.", "error");
    return;
  }
  if (selectedFile.size > MAX_FILE_SIZE) {
    setStatus(
      `The selected file is ${formatSize(
        selectedFile.size
      )}. Files larger than 20 MB are not supported.`,
      "error"
    );
    resetFileSelection();
    return;
  }

  const formData = new FormData();
  formData.append("file", selectedFile);
  const uomMode = uomSelect.value || "random";

  convertButton.disabled = true;
  convertButton.textContent = "Converting...";
  setStatus("Uploading and converting, please wait...");
  showSpinner();

  try {
    const response = await fetch(`/api/convert?uom_mode=${encodeURIComponent(
      uomMode
    )}`, {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      let message = "Conversion failed.";
      try {
        const payload = await response.json();
        if (payload?.detail) {
          if (typeof payload.detail === "string") {
            message = payload.detail;
          } else if (payload.detail?.detail) {
            message = payload.detail.detail;
          }
        }
      } catch (error) {
        // ignore JSON parse issues
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    const cd = response.headers.get("Content-Disposition") || "";
    const match = /filename="?([^"]+)"?/i.exec(cd);
    const filename = match ? match[1] : "k1-import.xlsx";

    const downloadUrl = URL.createObjectURL(blob);
    if (lastDownloadUrl) {
      URL.revokeObjectURL(lastDownloadUrl);
    }
    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();

    lastDownloadUrl = downloadUrl;
    lastFileLink.href = downloadUrl;
    lastFileLink.download = filename;
    lastFileLink.textContent = filename;
    lastFileSection.classList.remove("hidden");

    setStatus("Conversion succeeded. Download should begin shortly.", "success");
  } catch (error) {
    setStatus(error.message || "Conversion failed.", "error");
  } finally {
    convertButton.disabled = !selectedFile;
    convertButton.textContent = "Convert and Download";
    hideSpinner();
  }
});
