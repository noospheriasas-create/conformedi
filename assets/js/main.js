/* ============================================================
   noospheria — main.js
   ============================================================ */

(function () {
  "use strict";

  /* ---------- Loader ---------- */
  var loader = document.getElementById("loader");
  window.addEventListener("load", function () {
    setTimeout(function () {
      loader.classList.add("is-done");
      document.body.classList.add("is-loaded");
    }, 500);
  });
  // Fallback si l'événement load tarde
  setTimeout(function () {
    loader.classList.add("is-done");
    document.body.classList.add("is-loaded");
  }, 2500);

  /* ---------- Nav : fond au scroll + masquage à la descente ---------- */
  var nav = document.getElementById("nav");
  var lastY = 0;
  window.addEventListener(
    "scroll",
    function () {
      var y = window.scrollY;
      nav.classList.toggle("is-scrolled", y > 24);
      if (y > 400 && y > lastY) {
        nav.classList.add("is-hidden");
      } else {
        nav.classList.remove("is-hidden");
      }
      lastY = y;
    },
    { passive: true }
  );

  /* ---------- Menu mobile ---------- */
  var burger = document.getElementById("burger");
  var menu = document.getElementById("menu");
  burger.addEventListener("click", function () {
    var open = menu.classList.toggle("is-open");
    burger.classList.toggle("is-open", open);
    document.body.style.overflow = open ? "hidden" : "";
  });
  menu.querySelectorAll("a").forEach(function (link) {
    link.addEventListener("click", function () {
      menu.classList.remove("is-open");
      burger.classList.remove("is-open");
      document.body.style.overflow = "";
    });
  });

  /* ---------- Révélation au scroll ---------- */
  var revealObserver = new IntersectionObserver(
    function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          revealObserver.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.15, rootMargin: "0px 0px -40px 0px" }
  );
  document.querySelectorAll(".reveal").forEach(function (el) {
    revealObserver.observe(el);
  });

  /* ---------- Compteurs animés ---------- */
  function animateCount(el) {
    var target = parseInt(el.dataset.count, 10);
    var suffix = el.dataset.suffix || "";
    var duration = 1600;
    var start = null;

    function step(ts) {
      if (!start) start = ts;
      var progress = Math.min((ts - start) / duration, 1);
      // easeOutQuart
      var eased = 1 - Math.pow(1 - progress, 4);
      el.textContent = Math.round(eased * target) + suffix;
      if (progress < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }

  var countObserver = new IntersectionObserver(
    function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          animateCount(entry.target);
          countObserver.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.5 }
  );
  document.querySelectorAll(".stat__value").forEach(function (el) {
    countObserver.observe(el);
  });
})();
