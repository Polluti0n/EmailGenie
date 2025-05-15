
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Office.js is fully loaded and the add-in is ready
    initTheme();
    
    startLoading(() => {
      // Simulate loading work
      return new Promise((resolve) => setTimeout(resolve, 5000));
    });
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

const rand = (min, max) => Math.floor(Math.random() * (max - min) + min);

function startLoading(callback) {
  const docBody = document.body
  const loader = document.createElement("div");
  loader.classList.add("loader", "loader-18");

  for (let i = 1; i <= 8; i++) {
    const star = document.createElement("div");
    star.classList.add("css-star", `star${i}`);
    loader.append(star);
  }

  docBody.append(loader);

  // Execute the callback (can be async)
  Promise.resolve(callback()).finally(() => {
    triggerPoof(loader);
  });
}

function triggerPoof(loader) {
  const poof = document.createElement("span");
  poof.style = `
    position: absolute;
    top: 50%;
    left: 50%;
    translate: -50% -50%;
    width: ${loader.offsetWidth * 2}px;
    height: ${loader.offsetHeight * 2}px;
  `;
  loader.parentElement.append(poof);

  const shapes = [
    "https://static.vecteezy.com/system/resources/thumbnails/048/043/997/small/light-effect-ray-glow-sparkling-free-png.png",
    "https://static.vecteezy.com/system/resources/thumbnails/012/595/143/small/realistic-white-cloud-free-png.png",
    "https://static.vecteezy.com/system/resources/thumbnails/054/591/619/small/exploding-white-smoke-cloud-free-png.png"
  ];

  let parts = [];
  let animations = [];

  for (let y = 0; y < 4; y++) {
    for (let x = 0; x < 4; x++) {
      const part = document.createElement("span");
      part.style = `
        --cloud: url(${shapes[rand(0, shapes.length)]});
        background: linear-gradient(var(--tblr-theme-primary), var(--tblr-theme-primary)) no-repeat center / contain, var(--cloud) no-repeat center / contain;
        mask: var(--cloud) no-repeat center / contain;
        rotate: ${rand(0, 360)}deg;
      `;
      poof.append(part);
      parts.push(part);
    }
  }

  parts.forEach((p) => {
    animations.push(
      p.animate(
        [
          {
            transform: `translate(${rand(-400, 400)}px, ${rand(-400, 400)}px) scale(4)`,
            opacity: 0
          }
        ],
        {
          duration: 1000,
          easing: "ease-out",
          fill: "forwards",
          delay: rand(0, 100)
        }
      )
    );
  });

  animations.unshift(
    loader.animate([{ transform: "scale(2)", opacity: 0 }], {
      duration: 250,
      fill: "forwards"
    })
  );

  Promise.all(animations.map(a => a.finished)).then(() => {
    poof.remove();
    loader.remove();
  });
}
