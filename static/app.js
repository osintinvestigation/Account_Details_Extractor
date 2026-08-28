/* ============================================================
   Email Investigation Report Management System
   Frontend logic + 3D animations
   ============================================================ */

(() => {
  "use strict";

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);

  const form = $("upload-form");
  const fileInput = $("zipfile");
  const fileInfo = $("file-info");
  const dropZone = $("drop-zone");
  const submitBtn = $("submit-btn");

  const progressCard = $("progress-card");
  const progressBar = $("progress-bar");
  const progressStage = $("progress-stage");
  const progressCounter = $("progress-counter");
  const statusPill = $("status-pill");
  const logOutput = $("log-output");

  const resultCard = $("result-card");
  const statsGrid = $("stats-grid");
  const downloadGrid = $("download-grid");
  const newJobBtn = $("new-job-btn");

  const errorCard = $("error-card");
  const errorText = $("error-text");
  const errorLog = $("error-log");
  const errorDismiss = $("error-dismiss");

  // ---------- State ----------
  let currentJobId = null;
  let pollTimer = null;
  let lastLogCount = 0;

  // ---------- Helpers ----------
  const fmtBytes = (n) => {
    if (!n || n <= 0) return "0 B";
    const units = ["B", "KB", "MB", "GB"];
    let i = 0;
    let v = n;
    while (v >= 1024 && i < units.length - 1) {
      v /= 1024;
      i++;
    }
    return `${v.toFixed(v >= 10 || i === 0 ? 0 : 1)} ${units[i]}`;
  };

  const show = (el) => el.classList.remove("hidden");
  const hide = (el) => el.classList.add("hidden");

  const setStatus = (text, kind) => {
    statusPill.textContent = text;
    statusPill.classList.remove("processing", "completed", "failed");
    if (kind) statusPill.classList.add(kind);
  };

  const setStage = (text) => {
    progressStage.textContent = text;
  };

  const setProgress = (pct, label, indeterminate = false) => {
    if (indeterminate) {
      progressBar.classList.add("indeterminate");
    } else {
      progressBar.classList.remove("indeterminate");
      progressBar.style.width = `${Math.min(100, Math.max(0, pct))}%`;
    }
    if (label) progressCounter.textContent = label;
  };

  const stageFor = (line) => {
    if (!line) return null;
    const l = line.toLowerCase();
    if (l.includes("extracting main zip")) return { pct: 8, label: "Extracting ZIP" };
    if (l.includes("recursively extracting nested")) return { pct: 18, label: "Scanning nested ZIPs" };
    if (l.includes("nested zips extracted:")) return { pct: 26, label: "Nested ZIPs extracted" };
    if (l.includes("reading google data")) return { pct: 32, label: "Parsing HTML / CSV" };
    if (l.includes("extraction summary")) return { pct: 40, label: "Extraction summary" };
    if (l.includes("converting utc")) return { pct: 48, label: "Converting timestamps" };
    if (l.includes("creating excel report")) return { pct: 60, label: "Building Excel report" };
    if (l.includes("creating word report")) return { pct: 72, label: "Building Word report" };
    if (l.includes("creating formatted ip csv")) return { pct: 82, label: "Building IP CSV" };
    if (l.includes("creating separate ip word files")) return { pct: 90, label: "Building separate IP Word files" };
    if (l.includes("creating ip text files")) return { pct: 96, label: "Building IP text files" };
    if (l.includes("report generation completed")) return { pct: 100, label: "Done" };
    return null;
  };

  const appendLog = (lines) => {
    if (!Array.isArray(lines) || !lines.length) return;
    const newLines = lines.slice(lastLogCount);
    lastLogCount = lines.length;
    if (!newLines.length) return;
    const text = newLines.map((l) => l).join("\n") + "\n";
    logOutput.textContent += text;
    logOutput.scrollTop = logOutput.scrollHeight;
    const latest = newLines[newLines.length - 1];
    const s = stageFor(latest);
    if (s) setProgress(s.pct, s.label);
  };

  // ---------- Animated counter ----------
  const animateCounter = (el, target) => {
    const final = Number(target) || 0;
    const duration = 900;
    const start = performance.now();
    const startVal = 0;
    const tick = (now) => {
      const t = Math.min(1, (now - start) / duration);
      // easeOutCubic
      const eased = 1 - Math.pow(1 - t, 3);
      const current = Math.round(startVal + (final - startVal) * eased);
      el.textContent = current.toLocaleString();
      if (t < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  };

  // ---------- 3D tilt (mouse-driven perspective) ----------
  const attach3DTilt = (el, maxDeg = 6) => {
    let raf = null;
    el.addEventListener("mousemove", (e) => {
      const rect = el.getBoundingClientRect();
      const x = (e.clientX - rect.left) / rect.width - 0.5;
      const y = (e.clientY - rect.top) / rect.height - 0.5;
      if (raf) cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        el.style.transform =
          `perspective(900px) rotateX(${(-y * maxDeg).toFixed(2)}deg) ` +
          `rotateY(${(x * maxDeg).toFixed(2)}deg) translateZ(0)`;
      });
    });
    el.addEventListener("mouseleave", () => {
      if (raf) cancelAnimationFrame(raf);
      el.style.transform = "";
    });
  };

  // ---------- File input ----------
  const updateFileInfo = () => {
    const f = fileInput.files && fileInput.files[0];
    if (!f) {
      fileInfo.textContent = "No file selected";
      fileInfo.classList.remove("has-file");
      return;
    }
    fileInfo.textContent = `${f.name} (${fmtBytes(f.size)})`;
    fileInfo.classList.add("has-file");
  };

  fileInput.addEventListener("change", updateFileInfo);

  // Drag and drop.
  ["dragenter", "dragover"].forEach((evt) =>
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add("dragover");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("dragover");
    })
  );
  dropZone.addEventListener("drop", (e) => {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!f.name.toLowerCase().endsWith(".zip")) {
      showError("Only .zip files are accepted.");
      return;
    }
    const dt = new DataTransfer();
    dt.items.add(f);
    fileInput.files = dt.files;
    updateFileInfo();
  });

  // ---------- Submit ----------
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    hide(errorCard);

    const f = fileInput.files && fileInput.files[0];
    if (!f) {
      showError("Please select a ZIP file before generating the report.");
      return;
    }
    if (!f.name.toLowerCase().endsWith(".zip")) {
      showError("Only .zip files are accepted.");
      return;
    }

    startJob(f);
  });

  const startJob = async (file) => {
    submitBtn.disabled = true;
    submitBtn.classList.add("loading");
    submitBtn.querySelector(".btn-label").textContent = "Uploading…";

    hide(resultCard);
    show(progressCard);
    logOutput.textContent = "";
    lastLogCount = 0;
    setStatus("uploading", "processing");
    setStage("Uploading ZIP…");
    setProgress(0, "0 / 0", true);
    statsGrid.innerHTML = "";

    if (downloadGrid) {
      downloadGrid.querySelectorAll(".btn.download").forEach((btn) => {
        btn.disabled = true;
        btn.style.opacity = "";
        const ft = btn.dataset.fileType;
        const subtitleEl = btn.querySelector(
          `small[data-file-type="${ft}"]`
        );
        if (subtitleEl && btn.dataset.defaultSubtitle === undefined) {
          btn.dataset.defaultSubtitle = subtitleEl.textContent;
        }
        if (subtitleEl && btn.dataset.defaultSubtitle !== undefined) {
          subtitleEl.textContent = btn.dataset.defaultSubtitle;
        }
      });
    }

    const formData = new FormData();
    formData.append("zipfile", file);

    let resp;
    try {
      resp = await fetch("/upload", { method: "POST", body: formData });
    } catch (err) {
      submitBtn.disabled = false;
      submitBtn.classList.remove("loading");
      submitBtn.querySelector(".btn-label").textContent = "Generate Report";
      showError(`Network error: ${err.message || err}`);
      return;
    }

    if (!resp.ok) {
      let msg = `Upload failed (HTTP ${resp.status}).`;
      try {
        const data = await resp.json();
        if (data && data.error) msg = data.error;
      } catch (_) { /* ignore */ }
      submitBtn.disabled = false;
      submitBtn.classList.remove("loading");
      submitBtn.querySelector(".btn-label").textContent = "Generate Report";
      showError(msg);
      return;
    }

    let data;
    try {
      data = await resp.json();
    } catch (_) {
      showError("Server returned an invalid response.");
      return;
    }

    if (!data.job_id) {
      showError("Server did not return a job id.");
      return;
    }

    currentJobId = data.job_id;
    submitBtn.querySelector(".btn-label").textContent = "Processing…";
    setStage("Job queued");
    pollJob();
  };

  // ---------- Polling ----------
  const pollJob = () => {
    if (!currentJobId) return;
    if (pollTimer) clearTimeout(pollTimer);

    pollTimer = setTimeout(async () => {
      try {
        const resp = await fetch(`/status/${currentJobId}`);
        if (!resp.ok) {
          showError(`Status check failed (HTTP ${resp.status}).`);
          finishJobUI(false);
          return;
        }
        const data = await resp.json();
        handleStatus(data);
      } catch (err) {
        showError(`Lost connection while polling: ${err.message || err}`);
        finishJobUI(false);
      }
    }, 1200);
  };

  const handleStatus = (data) => {
    if (data.status === "pending") {
      setStatus("queued", "processing");
      setStage("Queued for processing");
      pollJob();
      return;
    }

    if (data.status === "processing") {
      setStatus("processing", "processing");
      appendLog(data.logs || []);
      pollJob();
      return;
    }

    if (data.status === "completed") {
      appendLog(data.logs || []);
      setStatus("completed", "completed");
      setStage("Completed");
      setProgress(100, "Done");
      progressBar.classList.remove("indeterminate");
      renderResults(data);
      finishJobUI(true);
      return;
    }

    if (data.status === "failed") {
      appendLog(data.logs || []);
      setStatus("failed", "failed");
      setStage("Failed");
      progressBar.classList.remove("indeterminate");
      showError(data.error || "Processing failed.", (data.logs || []).join("\n"));
      finishJobUI(false);
      return;
    }

    pollJob();
  };

  const finishJobUI = (success) => {
    submitBtn.disabled = false;
    submitBtn.classList.remove("loading");
    submitBtn.querySelector(".btn-label").textContent = "Generate Report";
  };

  // ---------- Results ----------
  const renderResults = (data) => {
    show(resultCard);
    statsGrid.innerHTML = "";

    const stats = data.stats || {};
    const tiles = [
      { label: "Google IP Activity", value: stats.total_ip, accent: true },
      { label: "Successful Logins", value: stats.successful },
      { label: "Failed Logins", value: stats.failed },
      { label: "Unique Google IPs", value: stats.unique_google_ips },
      { label: "Formatted IP Details", value: stats.formatted_ip_details },
      { label: "2405 / 2409 IPv6", value: stats.special_count },
      { label: "Remaining IPs", value: stats.remaining_count },
      { label: "Geolocated IPs", value: stats.geolocated_count },
      { label: "Phone Associations", value: stats.assoc_count },
      { label: "Connected", value: stats.connected_count },
      { label: "Devices", value: stats.device_count },
    ];

    tiles.forEach((t, i) => {
      const el = document.createElement("div");
      el.className = "stat-tile" + (t.accent ? " accent" : "");
      el.style.animationDelay = `${i * 60}ms`;
      el.innerHTML = `
        <span class="label">${t.label}</span>
        <span class="value">0</span>
      `;
      statsGrid.appendChild(el);

      const valueEl = el.querySelector(".value");
      const final = Number(t.value) || 0;
      setTimeout(() => animateCounter(valueEl, final), 220 + i * 60);

      attach3DTilt(el, 5);
    });

    // Update each download button's subtitle with the real file size.
    const r = data.result || {};
    Object.keys(r).forEach((fileType, i) => {
      const meta = r[fileType];
      const btn = downloadGrid.querySelector(
        `.btn.download[data-file-type="${fileType}"]`
      );
      if (!btn) return;
      const subtitleEl = btn.querySelector(
        `small[data-file-type="${fileType}"]`
      );
      if (subtitleEl) {
        subtitleEl.textContent = `${meta.subtitle} • ${fmtBytes(meta.size)}`;
      }
      btn.disabled = false;
      btn.style.opacity = "";
      btn.style.animationDelay = `${(tiles.length * 60) + 120 + i * 50}ms`;
      attach3DTilt(btn, 4);
    });

    // Disable buttons for any output that wasn't produced.
    downloadGrid.querySelectorAll(".btn.download").forEach((btn) => {
      const ft = btn.dataset.fileType;
      if (!r[ft]) {
        btn.disabled = true;
        btn.style.opacity = "0.45";
        const subtitleEl = btn.querySelector(`small[data-file-type="${ft}"]`);
        if (subtitleEl) subtitleEl.textContent = "Not produced for this export";
      }
    });

    resultCard.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const triggerDownload = (fileType) => {
    if (!currentJobId) return;
    const url = `/download/${currentJobId}/${fileType}`;
    const a = document.createElement("a");
    a.href = url;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
  };

  downloadGrid.addEventListener("click", (e) => {
    const btn = e.target.closest(".btn.download");
    if (!btn) return;
    const fileType = btn.dataset.fileType;
    if (fileType) triggerDownload(fileType);
  });

  newJobBtn.addEventListener("click", () => {
    if (pollTimer) clearTimeout(pollTimer);
    currentJobId = null;
    lastLogCount = 0;
    form.reset();
    updateFileInfo();
    logOutput.textContent = "";
    hide(progressCard);
    hide(resultCard);
    hide(errorCard);
    document.querySelector(".app-header").scrollIntoView({ behavior: "smooth" });
  });

  // ---------- Errors ----------
  const showError = (message, log = "") => {
    show(errorCard);
    errorText.textContent = message;
    errorLog.textContent = log || "";
    errorCard.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  errorDismiss.addEventListener("click", () => {
    hide(errorCard);
    document.querySelector(".app-header").scrollIntoView({ behavior: "smooth" });
  });

  // ---------- Card subtle tilt on hover ----------
  document.querySelectorAll(".card").forEach((c) => attach3DTilt(c, 2));
})();