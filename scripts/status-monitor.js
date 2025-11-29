(function () {
  "use strict";

  let statusInterval = null;

  window.startStatusMonitor = function () {
    if (statusInterval) return;

    statusInterval = setInterval(function () {
      if (window.goCheckStatus) {
        try {
          const status = window.goCheckStatus();
          updateStatusUI(status);
        } catch (err) {
          console.error("[STATUS] Error checking status:", err);
        }
      }
    }, 250); // Poll every 250ms
  };

  window.stopStatusMonitor = function () {
    if (statusInterval) {
      clearInterval(statusInterval);
      statusInterval = null;
    }
  };

  function updateStatusUI(status) {
    if (!status) return;

    const statusText = document.getElementById("statusText");
    const currentFit = document.getElementById("currentFit");
    const currentParams = document.getElementById("currentParams");
    const startBtn = document.getElementById("startBtn");

    // Update status text
    const statusStr = status.current_status || status.CurrentStatus || "unknown";
    statusText.textContent = capitalize(statusStr);

    // Update status class
    statusText.className = "status-value";
    if (statusStr === "running" || statusStr === "fitting") {
      statusText.classList.add("running");
    } else if (statusStr === "ready" || statusStr === "done" || statusStr === "complete") {
      statusText.classList.add("ready");
    } else if (statusStr === "error") {
      statusText.classList.add("error");
    }

    // Update best fit
    const fit = status.current_best_fit || status.CurrentBestFit;
    if (fit !== undefined && fit !== null && fit < 1e308) {
      currentFit.textContent = formatNumber(fit);
    } else {
      currentFit.textContent = "—";
    }

    // Update parameters
    const params = status.current_paramters || status.CurrentParamters || status.current_parameters || [];
    if (params && params.length > 0) {
      currentParams.textContent = params.map(p => formatNumber(p)).join(", ");
    } else {
      currentParams.textContent = "—";
    }

    // Enable/disable start button based on status
    if (statusStr === "running" || statusStr === "fitting") {
      startBtn.disabled = true;
    } else {
      startBtn.disabled = false;
    }
  }

  function capitalize(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  function formatNumber(num) {
    if (typeof num !== "number") return num;
    if (Math.abs(num) < 0.0001 || Math.abs(num) > 10000) {
      return num.toExponential(4);
    }
    return num.toFixed(4);
  }
})();