// ===================== Кушодан/пӯшидани менюи мобилӣ =====================
document.addEventListener("DOMContentLoaded", function () {
  var toggle = document.getElementById("nav-toggle");
  var nav = document.getElementById("main-nav");

  if (toggle && nav) {
    toggle.addEventListener("click", function () {
      nav.classList.toggle("open");
    });
  }

  // ===================== Аккордеони дарсҳо =====================
  var lessonHeaders = document.querySelectorAll(".lesson-header");

  lessonHeaders.forEach(function (header) {
    header.addEventListener("click", function () {
      var lesson = header.closest(".lesson");
      var body = lesson.querySelector(".lesson-body");

      if (lesson.classList.contains("open")) {
        lesson.classList.remove("open");
        body.style.maxHeight = null;
      } else {
        lesson.classList.add("open");
        body.style.maxHeight = body.scrollHeight + "px";
      }
    });
  });

  // Агар саҳифа бо #lesson-... кушода шавад, мавзуи мувофиқро кушо
  if (window.location.hash) {
    var target = document.querySelector(window.location.hash);
    if (target && target.classList.contains("lesson")) {
      target.classList.add("open");
      var body = target.querySelector(".lesson-body");
      if (body) {
        body.style.maxHeight = body.scrollHeight + "px";
      }
      setTimeout(function () {
        target.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 100);
    }
  }

  // ===================== Ҷустуҷӯ дар луғат =====================
  var searchInput = document.getElementById("glossary-search");
  if (searchInput) {
    var items = document.querySelectorAll(".glossary-item");
    var emptyMsg = document.querySelector(".glossary-empty");

    searchInput.addEventListener("input", function () {
      var query = searchInput.value.trim().toLowerCase();
      var visibleCount = 0;

      items.forEach(function (item) {
        var text = item.textContent.toLowerCase();
        if (text.indexOf(query) !== -1) {
          item.style.display = "";
          visibleCount++;
        } else {
          item.style.display = "none";
        }
      });

      if (emptyMsg) {
        emptyMsg.style.display = visibleCount === 0 ? "block" : "none";
      }
    });
  }
});
