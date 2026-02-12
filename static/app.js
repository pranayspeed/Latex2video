const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const form = document.getElementById("uploadForm");
const statusEl = document.getElementById("status");
const logBox = document.getElementById("logBox");
const videoControls = document.querySelector(".video-controls");
const modeInputs = document.querySelectorAll('input[name="mode"]');

let pollTimer = null;

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b42318" : "#4a4a4a";
}

function updateFileName() {
  if (fileInput.files.length > 0) {
    fileName.textContent = fileInput.files[0].name;
  } else {
    fileName.textContent = "";
  }
}

function updateVideoControlsVisibility() {
  const selected = document.querySelector('input[name="mode"]:checked');
  if (!videoControls) return;
  videoControls.style.display = selected && selected.value === "video" ? "grid" : "none";
}

["dragenter", "dragover"].forEach(event => {
  dropzone.addEventListener(event, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add("dragover");
  });
});

dropzone.addEventListener("dragleave", (e) => {
  e.preventDefault();
  e.stopPropagation();
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  e.stopPropagation();
  dropzone.classList.remove("dragover");
  if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
    fileInput.files = e.dataTransfer.files;
    updateFileName();
  }
});

fileInput.addEventListener("change", updateFileName);
modeInputs.forEach((input) => input.addEventListener("change", updateVideoControlsVisibility));
updateVideoControlsVisibility();

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (pollTimer) {
    window.clearInterval(pollTimer);
    pollTimer = null;
  }
  if (!fileInput.files.length) {
    setStatus("Please choose a zip file.", true);
    return;
  }

  const data = new FormData(form);
  setStatus("Uploading and converting. This can take a minute...");
  logBox.textContent = "";

  try {
    const res = await fetch("/convert", {
      method: "POST",
      body: data,
    });

    if (!res.ok) {
      const payload = await res.json().catch(() => ({}));
      throw new Error(payload.error || "Conversion failed.");
    }

    const payload = await res.json();
    if (!payload.job_id) {
      throw new Error("Missing job id from server.");
    }
    const jobId = payload.job_id;

    pollTimer = window.setInterval(async () => {
      try {
        const statusRes = await fetch(`/status/${jobId}`);
        const statusPayload = await statusRes.json();
        if (!statusRes.ok) {
          throw new Error(statusPayload.error || "Failed to fetch job status.");
        }

        logBox.textContent = statusPayload.log || "";
        logBox.scrollTop = logBox.scrollHeight;

        if (statusPayload.status === "error") {
          window.clearInterval(pollTimer);
          pollTimer = null;
          setStatus(statusPayload.error || "Conversion failed.", true);
          return;
        }

        if (statusPayload.status === "done" && statusPayload.ready) {
          window.clearInterval(pollTimer);
          pollTimer = null;
          setStatus("Done. Downloading output...");
          const a = document.createElement("a");
          a.href = `/download/${jobId}`;
          a.download = "";
          document.body.appendChild(a);
          a.click();
          a.remove();
          setStatus("Done. Your file should be downloaded.");
        }
      } catch (pollErr) {
        if (pollTimer) {
          window.clearInterval(pollTimer);
          pollTimer = null;
        }
        setStatus(pollErr.message, true);
      }
    }, 1000);
  } catch (err) {
    if (pollTimer) {
      window.clearInterval(pollTimer);
      pollTimer = null;
    }
    setStatus(err.message, true);
  }
});
