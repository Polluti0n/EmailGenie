
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Office.js is fully loaded and the add-in is ready
    initTheme();
  }
});
const themeConfig = {
	"theme": "light",
	"theme-base": "gray",
	"theme-font": "sans-serif",
	"theme-primary": "blue",
	"theme-radius": "1",
}

function initTheme() {
  const themeConfig = {
    "theme": "light",
    "theme-base": "gray",
    "theme-font": "sans-serif",
    "theme-primary": "blue",
    "theme-radius": "1",
  };

  const params = new Proxy(new URLSearchParams(window.location.search), {
    get: (searchParams, prop) => searchParams.get(prop),
  });

  // Load existing roaming settings or start with an empty object
  let storedTheme = Office.context.roamingSettings.get("theme");
  if (typeof storedTheme !== "object" || storedTheme === null) {
    storedTheme = {};
  }

  // Track changes to save them later
  const updatedTheme = {};

  for (const key in themeConfig) {
    const param = params[key];
    let selectedValue;

    if (param) {
      selectedValue = param;
    } else if (storedTheme[key]) {
      selectedValue = storedTheme[key];
    } else {
      const localValue = localStorage.getItem("tabler-" + key);
      selectedValue = localValue || themeConfig[key];
    }

    // Apply to document
    if (selectedValue !== themeConfig[key]) {
      document.documentElement.setAttribute("data-bs-" + key, selectedValue);
    } else {
      document.documentElement.removeAttribute("data-bs-" + key);
    }

    // Persist to localStorage and prepare to save to roaming settings
    localStorage.setItem("tabler-" + key, selectedValue);
    updatedTheme[key] = selectedValue;
  }

  // Save updated theme to roaming settings
  Office.context.roamingSettings.set("theme", updatedTheme);
  Office.context.roamingSettings.saveAsync();
}