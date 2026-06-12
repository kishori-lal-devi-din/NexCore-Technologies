const header = document.getElementById("siteHeader");
const menuToggle = document.getElementById("menuToggle");
const mobilePanel = document.getElementById("mobilePanel");
const uptimeCounter = document.getElementById("uptimeCounter");
const uptimeStart = new Date("2024-02-12T09:58:00Z");

function setHeaderState() {
  header.classList.toggle("is-scrolled", window.scrollY > 12);
}

function setMenu(open) {
  header.classList.toggle("is-open", open);
  menuToggle.setAttribute("aria-expanded", String(open));
  menuToggle.setAttribute("aria-label", open ? "Close navigation menu" : "Open navigation menu");
}

function updateUptimeCounter() {
  if (!uptimeCounter) return;
  const elapsed = Math.max(0, Date.now() - uptimeStart.getTime());
  const totalMinutes = Math.floor(elapsed / 60000);
  const days = Math.floor(totalMinutes / 1440);
  const hours = Math.floor((totalMinutes % 1440) / 60);
  const minutes = totalMinutes % 60;
  uptimeCounter.textContent = `99.98% Uptime - ${days}d ${hours}h ${minutes}m`;
}

setHeaderState();
updateUptimeCounter();
window.addEventListener("scroll", setHeaderState, { passive: true });
window.addEventListener("resize", () => {
  if (window.innerWidth >= 1024) setMenu(false);
});

menuToggle.addEventListener("click", () => {
  setMenu(!header.classList.contains("is-open"));
});

mobilePanel.addEventListener("click", (event) => {
  if (event.target.closest("a")) setMenu(false);
});

if (uptimeCounter) {
  setInterval(updateUptimeCounter, 60000);
}

const contactForm = document.getElementById("contactForm");
const contactService = document.getElementById("contactService");
const contactStatus = document.getElementById("contactStatus");
const serviceOptions = document.querySelectorAll(".service-option");
const googleSheetsPlaceholder = "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL";

function updateSelectedServices() {
  if (!contactService) return;
  const selected = Array.from(serviceOptions)
    .filter((item) => item.classList.contains("is-selected"))
    .map((item) => item.dataset.service);
  contactService.value = selected.join(", ");
}

serviceOptions.forEach((option) => {
  option.addEventListener("click", () => {
    option.classList.toggle("is-selected");
    const selectedCount = Array.from(serviceOptions).filter((item) =>
      item.classList.contains("is-selected")
    ).length;
    if (selectedCount === 0) option.classList.add("is-selected");
    updateSelectedServices();
  });
});

updateSelectedServices();

function resetContactForm() {
  contactForm.reset();
  serviceOptions.forEach((item) => item.classList.remove("is-selected"));
  if (serviceOptions[0]) serviceOptions[0].classList.add("is-selected");
  updateSelectedServices();
}

if (contactForm && contactStatus) {
  contactForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const submitBtn = contactForm.querySelector('button[type="submit"]');
    const originalLabel = submitBtn.textContent;
    const endpoint = contactForm.action.trim();

    if (!endpoint || endpoint.includes(googleSheetsPlaceholder)) {
      contactStatus.textContent =
        "Google Sheets is not connected yet. Add your deployed Apps Script Web App URL to the form action.";
      contactStatus.className = "status-error";
      return;
    }

    const formData = new FormData(contactForm);
    const payload = Object.fromEntries(formData.entries());
    payload["Source Page"] = window.location.href;

    submitBtn.disabled = true;
    submitBtn.setAttribute("aria-busy", "true");
    submitBtn.textContent = "Sending...";
    contactStatus.textContent = "";
    contactStatus.className = "";

    try {
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        body: JSON.stringify(payload),
      });

      contactStatus.textContent =
        "Thanks! Your request has been submitted. We will follow up within 24 hours.";
      contactStatus.classList.add("status-success");
      resetContactForm();
    } catch {
      contactStatus.textContent =
        "Submission failed. Please try again or email consulting@nexcoretechnologies.com.";
      contactStatus.classList.add("status-error");
    } finally {
      submitBtn.disabled = false;
      submitBtn.removeAttribute("aria-busy");
      submitBtn.textContent = originalLabel;
    }
  });
}
