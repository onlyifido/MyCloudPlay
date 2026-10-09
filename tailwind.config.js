module.exports = {
  content: ["./index.html", "./app.js", "./appearance.js"],
  darkMode: "class",
  theme: {
    extend: {
      fontFamily: { sans: ["Inter", "system-ui", "sans-serif"] },
      colors: {
        accent: "var(--accent)",
        "accent-soft": "var(--accent-soft)",
        "accent-border": "var(--accent-border)",
        "accent-contrast": "var(--accent-contrast)",
        gray: { 900: "#121212", 950: "#0a0a0a" },
      },
    },
  },
};
