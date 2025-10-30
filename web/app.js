// DOM element selections
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const pickBtn = document.getElementById("pick-btn");
const fileInfo = document.getElementById("file-info");
const convertForm = document.getElementById("convert-form");
const convertButton = document.getElementById("convert-button");
const statusEl = document.getElementById("status");
const lastFileSection = document.getElementById("last-file");
const lastFileLink = document.getElementById("last-file-link");
const countrySelect = document.getElementById("country-select");

// State variables
let selectedFile = null;
let lastDownloadUrl = null;

// File validation constants
const ACCEPTED_TYPES = [
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
];
const ACCEPTED_EXTENSIONS = [".xls", ".xlsx"];
const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20 MB

// Create and inject a spinner element for loading states.
const spinner = document.createElement("div");
spinner.className = "spinner hidden";
spinner.setAttribute("role", "status");
spinner.setAttribute("aria-live", "polite");
spinner.innerHTML =
  '<span class="spinner-icon" aria-hidden="true"></span><span class="spinner-text">Converting…</span>';
convertForm.insertAdjacentElement("afterend", spinner);

// Spinner visibility functions
const showSpinner = () => spinner.classList.remove("hidden");
const hideSpinner = () => spinner.classList.add("hidden");
let picking = false; // Prevents multiple file pick dialogs

/**
 * Resets the status message.
 */
const resetStatus = () => {
  statusEl.textContent = "";
  statusEl.classList.remove("error", "success");
};

/**
 * Sets the status message with a specified type (info, error, success).
 * @param {string} message - The message to display.
 * @param {string} type - The type of message ('info', 'error', 'success').
 */
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

/**
 * Formats a file size in bytes into a human-readable string (KB, MB).
 * @param {number} bytes - The file size in bytes.
 * @returns {string} The formatted file size.
 */
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

/**
 * Checks if a file is a valid Excel file based on its MIME type or extension.
 * @param {File} file - The file to validate.
 * @returns {boolean} True if the file is valid, false otherwise.
 */
const isValidFile = (file) => {
  if (!file) return false;
  const hasValidType = ACCEPTED_TYPES.includes(file.type);
  const hasValidExtension = ACCEPTED_EXTENSIONS.some((ext) =>
    file.name.toLowerCase().endsWith(ext)
  );
  return hasValidType || hasValidExtension;
};

/**
 * Updates the UI to display information about the selected file.
 */
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

/**
 * Resets the file selection state and optionally displays an error message.
 * @param {string} [message] - An optional error message to display.
 */
const resetFileSelection = (message) => {
  selectedFile = null;
  fileInput.value = "";
  updateFileInfo();
  if (message) {
    setStatus(message, "error");
  }
};

/**
 * Handles the selection of files, validates them, and updates the UI.
 * @param {FileList} files - The list of selected files.
 */
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

// Drag and drop event handlers
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

// Attach drag and drop event listeners to the dropzone.
dropzone.addEventListener("dragenter", onDragEnter);
dropzone.addEventListener("dragover", onDragOver);
dropzone.addEventListener("dragleave", onDragLeave);
dropzone.addEventListener("drop", onDrop);

// Event listener for the "Pick a file" button.
pickBtn.addEventListener("click", () => {
  if (picking) return;
  picking = true;
  console.log("pickBtn clicked");
  fileInput.click();
  setTimeout(() => {
    picking = false;
  }, 500);
});

// Event listener for the file input element.
fileInput.addEventListener("change", (event) => {
  handleFiles(event.target.files);
});

// Make the entire dropzone clickable to open the file picker.
dropzone.addEventListener("click", () => pickBtn.click());

// Event listener for the conversion form submission.
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

  // Prepare the form data for the API request.
  const formData = new FormData();
  formData.append("file", selectedFile);
  const country = countrySelect.value || "ID";

  // Update UI to reflect the loading state.
  convertButton.disabled = true;
  convertButton.textContent = "Converting...";
  setStatus("Uploading and converting, please wait...");
  showSpinner();

  try {
    // Make the API call to the conversion endpoint.
    const response = await fetch(`/api/convert?country=${encodeURIComponent(
      country
    )}`, {
      method: "POST",
      body: formData,
    });

    // Handle non-successful responses.
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

    // Process the successful response.
    const blob = await response.blob();
    const cd = response.headers.get("Content-Disposition") || "";
    const match = /filename="?([^"]+)"?/i.exec(cd);
    const filename = match ? match[1] : "k1-import.xlsx";

    // Create a download link and trigger the download.
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

    // Update the "Last File" section with the new download link.
    lastDownloadUrl = downloadUrl;
    lastFileLink.href = downloadUrl;
    lastFileLink.download = filename;
    lastFileLink.textContent = filename;
    lastFileSection.classList.remove("hidden");

    setStatus("Conversion succeeded. Download should begin shortly.", "success");
  } catch (error) {
    // Handle errors during the fetch or conversion process.
    setStatus(error.message || "Conversion failed.", "error");
  } finally {
    // Reset the UI from the loading state.
    convertButton.disabled = !selectedFile;
    convertButton.textContent = "Convert and Download";
    hideSpinner();
  }
});
