// loader.js
const createLoader = () => {
  const rand = (min, max) => Math.floor(Math.random() * (max - min) + min);

  function triggerPoof(loader) {
    const container = document.querySelector('.loader-overlay');
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
        part.classList.add("loading-part");
        part.style = `
          --poof-shape: url(${shapes[rand(0, shapes.length)]});
          --poof-rotate: ${rand(0, 360)}deg;
        `;
        poof.append(part);
        parts.push(part);
      }
    }

    parts.forEach((p) => {
      let bounds = p.getBoundingClientRect();
      animations.push(
        p.animate(
          [
            {
              transform: `translate(${rand(bounds.left*-1, container.offsetWidth-(bounds.left+bounds.width))}px, ${rand(bounds.top*-1, container.offsetHeight-(bounds.top+bounds.height))}px) scale(6)`,
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

    animations.push(
      container.animate({ opacity: 0 }, { duration: 200 })
    );

    animations.unshift(
      loader.animate([{ transform: "scale(0)" }], {
        duration: 200,
        fill: "forwards"
      })
    );

    Promise.all(animations.map(a => a.finished)).then(() => {
      poof.remove();
      loader.remove();
      container.remove();
    });
  }

  function startLoading(callback) {
    const container = document.createElement("div");
    container.classList.add("loader-overlay");
    const loader = document.createElement("div");
    loader.classList.add("star-loader");

    for (let i = 1; i <= 8; i++) {
      const star = document.createElement("div");
      star.classList.add("css-star", `star${i}`);
      loader.append(star);
    }

    container.append(loader);
    document.body.append(container);

    container.animate({ opacity: 1 }, {
      duration: 200,
      fill: 'forwards'
    });

    Promise.resolve(callback()).finally(() => {
      triggerPoof(loader);
    });
  }

  return startLoading;
};

export default createLoader;
