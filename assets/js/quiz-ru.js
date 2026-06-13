// ===================== Логика теста =====================
document.addEventListener("DOMContentLoaded", function () {
  var container = document.getElementById("quiz-container");
  if (!container) return;

  var resultBox = document.getElementById("quiz-result");
  var checkBtn = document.getElementById("check-btn");
  var resetBtn = document.getElementById("reset-btn");
  var selectButtons = document.querySelectorAll(".quiz-select button");
  var currentQuestions = [];

  function shuffle(array) {
    var copy = array.slice();
    for (var i = copy.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var tmp = copy[i];
      copy[i] = copy[j];
      copy[j] = tmp;
    }
    return copy;
  }

  function renderQuestions(questions) {
    container.innerHTML = "";

    questions.forEach(function (q, qIndex) {
      var qDiv = document.createElement("div");
      qDiv.className = "quiz-question";
      qDiv.dataset.correct = q.correct;

      var qText = document.createElement("p");
      qText.className = "question-text";
      qText.textContent = (qIndex + 1) + ". " + q.question;
      qDiv.appendChild(qText);

      var optionsDiv = document.createElement("div");
      optionsDiv.className = "quiz-options";

      q.options.forEach(function (opt, oIndex) {
        var label = document.createElement("label");
        var input = document.createElement("input");
        input.type = "radio";
        input.name = "question-" + qIndex;
        input.value = oIndex;
        label.appendChild(input);
        label.appendChild(document.createTextNode(opt));
        optionsDiv.appendChild(label);
      });

      qDiv.appendChild(optionsDiv);

      var feedback = document.createElement("div");
      feedback.className = "answer-feedback";
      feedback.textContent = "Правильный ответ: " + q.options[q.correct] + ". " + q.explanation;
      qDiv.appendChild(feedback);

      container.appendChild(qDiv);
    });

    resultBox.style.display = "none";
    resultBox.textContent = "";
  }

  function startQuiz(category) {
    if (category === "mixed") {
      currentQuestions = shuffle(quizData.mikro.concat(quizData.makro));
    } else {
      currentQuestions = quizData[category];
    }
    renderQuestions(currentQuestions);
    container.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  selectButtons.forEach(function (btn) {
    btn.addEventListener("click", function () {
      selectButtons.forEach(function (b) {
        b.classList.remove("active");
      });
      btn.classList.add("active");
      startQuiz(btn.dataset.category);
    });
  });

  checkBtn.addEventListener("click", function () {
    var questionDivs = container.querySelectorAll(".quiz-question");
    var score = 0;
    var answered = 0;

    questionDivs.forEach(function (qDiv, qIndex) {
      var selected = qDiv.querySelector('input[name="question-' + qIndex + '"]:checked');
      var correctIndex = parseInt(qDiv.dataset.correct, 10);

      qDiv.classList.add("checked");
      qDiv.classList.remove("correct", "incorrect");

      if (selected) {
        answered++;
        if (parseInt(selected.value, 10) === correctIndex) {
          qDiv.classList.add("correct");
          score++;
        } else {
          qDiv.classList.add("incorrect");
        }
      } else {
        qDiv.classList.add("incorrect");
      }
    });

    var total = questionDivs.length;
    var percent = Math.round((score / total) * 100);
    var message = "Ваш результат: " + score + " из " + total + " (" + percent + "%)";

    if (answered < total) {
      message += " — некоторые вопросы остались без ответа.";
    } else if (percent === 100) {
      message += " — отлично! Все ответы правильные!";
    } else if (percent >= 60) {
      message += " — хороший результат!";
    } else {
      message += " — повторите тему ещё раз.";
    }

    resultBox.textContent = message;
    resultBox.style.display = "block";
    resultBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
  });

  resetBtn.addEventListener("click", function () {
    var active = document.querySelector(".quiz-select button.active");
    var category = active ? active.dataset.category : "mikro";
    startQuiz(category);
  });

  // Тест по микроэкономике открывается по умолчанию
  selectButtons[0].classList.add("active");
  startQuiz(selectButtons[0].dataset.category);
});
