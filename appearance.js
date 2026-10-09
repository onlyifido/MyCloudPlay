/* Appearance is applied before paint; unavailable browser storage is non-fatal. */
const storage = {
  getItem(key) {
    try {
      return localStorage.getItem(key);
    } catch {
      return null;
    }
  },
  setItem(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch {
      /* Keep the current session usable. */
    }
  },
  removeItem(key) {
    try {
      localStorage.removeItem(key);
    } catch {
      /* Storage may be disabled. */
    }
  },
};
function readStoredJSON(key, fallback, valid) {
  try {
    const raw = storage.getItem(key);
    if (raw === null) return fallback;
    const value = JSON.parse(raw);
    if (!valid(value)) throw new Error("Invalid stored data");
    return value;
  } catch {
    storage.removeItem(key);
    return fallback;
  }
}
const availableSkins = ["ocean", "moss", "sand", "rose", "graphite"];
const systemColorScheme = window.matchMedia("(prefers-color-scheme: dark)");
let appearanceMode = ["light", "dark", "system"].includes(
  storage.getItem("theme"),
)
  ? storage.getItem("theme")
  : "system";
let appearanceSkin = availableSkins.includes(storage.getItem("skin"))
  ? storage.getItem("skin")
  : "ocean";
function applyAppearance() {
  const dark =
    appearanceMode === "dark" ||
    (appearanceMode === "system" && systemColorScheme.matches);
  document.documentElement.classList.toggle("dark", dark);
  document.documentElement.classList.toggle("light", !dark);
  document.documentElement.dataset.skin = appearanceSkin;
  document.documentElement.style.colorScheme = dark ? "dark" : "light";
  document
    .querySelector('meta[name="theme-color"]')
    ?.setAttribute("content", dark ? "#131619" : "#f7f8fa");
  document
    .querySelectorAll("[data-mode]")
    .forEach((button) =>
      button.setAttribute(
        "aria-pressed",
        String(button.dataset.mode === appearanceMode),
      ),
    );
  document
    .querySelectorAll("button[data-skin]")
    .forEach((button) =>
      button.setAttribute(
        "aria-pressed",
        String(button.dataset.skin === appearanceSkin),
      ),
    );
}
function chooseMode(mode) {
  if (!["light", "dark", "system"].includes(mode)) return;
  appearanceMode = mode;
  storage.setItem("theme", mode);
  applyAppearance();
}
function chooseSkin(skin) {
  if (!availableSkins.includes(skin)) return;
  appearanceSkin = skin;
  storage.setItem("skin", skin);
  applyAppearance();
}
systemColorScheme.addEventListener("change", applyAppearance);
applyAppearance();
document.addEventListener("DOMContentLoaded", applyAppearance);
