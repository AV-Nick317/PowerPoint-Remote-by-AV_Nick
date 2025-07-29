// ============================================================
//  CENTRALIZED SLIDE STATE
// ============================================================

const SlideState = {
    number: 1,            // current slide number
    prefix: "date",       // z value from 0.txt
    firstTrigger: null,   // theme toggle trigger 1
    secondTrigger: null,  // theme toggle trigger 2
    noTimer: false,       // skip slide timer?
    timerDuration: 0,     // ms from first line of .txt
    activeTimer: null,    // auto-advance timeout
    countdownInterval: null // countdown display interval
};


// ============================================================
//  INITIALIZATION
// ============================================================

document.addEventListener("DOMContentLoaded", () => {
    setupThemeToggle();
    setupNavigationButtons();
    setupSlideForm();
    setupLegacyButton();

    loadSlideGallery("powerpoint_parts/0.txt").then(() => {
        loadMessage();

        setTimeout(() => {
            SlideState.number = 0;
            nextSlideImg();
        }, 1000);
    });
});


// ============================================================
//  SLIDE GALLERY
// ============================================================

async function loadSlideGallery(file) {
    const response = await fetch(file);
    const text = await response.text();
    const [countLine, prefixLine] = text.split("\n");

    const slideCount = Number(countLine);
    SlideState.prefix = prefixLine.trim();

    const gallery = document.getElementById("slideGallery");
    gallery.innerHTML = "";

    for (let i = 1; i <= slideCount; i++) {
        const btn = document.createElement("button");
        const img = document.createElement("img");

        img.src = `powerpoint_parts/${SlideState.prefix}${i}.jpg`;
        img.alt = "";

        btn.addEventListener("click", () => {
            JumpToSlide(i);
            toggleSlideGallery();
        });

        btn.appendChild(img);
        gallery.appendChild(btn);
    }
}

function toggleSlideGallery() {
    const div = document.getElementById("slideGallery");
    div.style.display = div.style.display === "none" ? "grid" : "none";
}


// ============================================================
//  THEME TOGGLING
// ============================================================

function setupThemeToggle() {
    const toggleBtn = document.getElementById("toggle-theme-btn");
    const themes = document.querySelectorAll("#mainTheme, #notesTheme");

    toggleBtn.addEventListener("click", () => {
        themes.forEach(theme => theme.disabled = !theme.disabled);
    });
}

function toggleStylesheets() {
    const main = document.getElementById("mainTheme");
    const notes = document.getElementById("notesTheme");

    const mainActive = !main.disabled;
    main.disabled = mainActive;
    notes.disabled = !mainActive;
}


// ============================================================
//  SLIDE MESSAGE + TIMER LOGIC
// ============================================================

async function loadMessage() {
    const file = `powerpoint_parts/${SlideState.prefix}${SlideState.number}.txt`;
    const response = await fetch(file);
    const text = await response.text();

    const lines = text.split("\n");
    SlideState.timerDuration = Number(lines[0]) * 1000;

    clearTimeout(SlideState.activeTimer);
    clearInterval(SlideState.countdownInterval);

    if (SlideState.timerDuration > 0 && !SlideState.noTimer) {
        startCountdown(SlideState.timerDuration);

        const slideAtStart = SlideState.number;
        SlideState.activeTimer = setTimeout(() => {
            if (SlideState.number === slideAtStart) nextSlideTimer();
        }, SlideState.timerDuration);
    } else {
        updateCountdownDisplay(0, 0);
    }

    const bodyText = lines.slice(1).join("\n");
    processCustomCommands(bodyText);

    document.getElementById("notesText").innerHTML = bodyText;
}

function processCustomCommands(text) {
    if (!text.startsWith("toggleRemoteTheme")) return;

    const match = text.match(/\(([^)]+)\)/);

    if (!match) {
        toggleStylesheets();
        return;
    }

    const numbers = match[1].match(/\d+/g).map(Number);

    SlideState.firstTrigger = SlideState.number + numbers[0];
    if (numbers[1]) {
        SlideState.secondTrigger = SlideState.firstTrigger + numbers[1];
    }
}

function startCountdown(duration) {
    const end = Date.now() + duration;

    SlideState.countdownInterval = setInterval(() => {
        const remaining = end - Date.now();
        if (remaining <= 0) {
            updateCountdownDisplay(0, duration / 1000);
            clearInterval(SlideState.countdownInterval);
            return;
        }

        const minutes = Math.floor((remaining % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((remaining % (1000 * 60)) / 1000);
        updateCountdownDisplay(minutes, seconds, duration / 1000);
    }, 1000);
}

function updateCountdownDisplay(min, sec, total = 0) {
    const padded = sec < 10 ? `0${sec}` : sec;
    document.getElementById("slideTimerCountdown").innerHTML = `${min}:${padded} / ${total}`;
}


// ============================================================
//  SLIDE IMAGE LOADING
// ============================================================

function updateSlideImages() {
    document.getElementById("picture").src =
        `powerpoint_parts/${SlideState.prefix}${SlideState.number}.jpg`;

    document.getElementById("picture2").src =
        `powerpoint_parts/${SlideState.prefix}${SlideState.number + 1}.jpg`;
}

function callSlideParts() {
    clearTimeout(SlideState.activeTimer);
    clearInterval(SlideState.countdownInterval);

    updateSlideImages();
    loadMessage();

    if (SlideState.firstTrigger === SlideState.number) {
        toggleStylesheets();
        SlideState.firstTrigger = null;
    }

    if (SlideState.secondTrigger === SlideState.number) {
        toggleStylesheets();
        SlideState.secondTrigger = null;
    }

    document.getElementById("inputBoxforSlide")
        .setAttribute("placeholder", SlideState.number);
}


// ============================================================
//  SLIDE NAVIGATION
// ============================================================

function nextSlideTimer() {
    SlideState.number++;
    SlideState.noTimer = false;
    callSlideParts();
}

function nextSlideImg() {
    SlideState.number++;
    SlideState.noTimer = true;
    callSlideParts();
}

function previousSlideImg() {
    SlideState.number--;
    SlideState.noTimer = true;
    callSlideParts();
}

function nextSlideAll() {
    nextPage();
    SlideState.number++;
    SlideState.noTimer = false;
    callSlideParts();
}

function previousSlideAll() {
    previousPage();
    SlideState.number--;
    SlideState.noTimer = true;
    callSlideParts();
}

function favSlideButton() {
    slideNumber(15);
    SlideState.number = 15;
    SlideState.noTimer = false;
    callSlideParts();
}

function JumpToSlide(n) {
    slideNumber(n);
    SlideState.number = n;
    SlideState.noTimer = false;
    callSlideParts();
}


// ============================================================
//  BUTTON + FORM SETUP
// ============================================================

function setupNavigationButtons() {
    const bind = (id, fn) =>
        document.getElementById(id).addEventListener("click", fn);

    bind("go-fav-slide", favSlideButton);
    bind("toggle-slides-btn", toggleSlideGallery);

    bind("go-previousSlideAll", previousSlideAll);
    bind("go-previousSlideImg", previousSlideImg);
    bind("go-previousPage", previousPage);

    bind("go-nextPage", nextPage);
    bind("go-nextSlideImg", nextSlideImg);
    bind("go-nextSlideAll", nextSlideAll);

    bind("picture", nextSlideAll);
    bind("picture2", nextSlideAll);
}

function setupSlideForm() {
    const form = document.getElementById("slideForm");

    form.addEventListener("submit", e => {
        e.preventDefault();
        const input = form.querySelector('input[name="slideSelect"]');

        if (!input.value.trim()) {
            alert("Slide Number is required.");
            return;
        }

        SlideState.number = Number(input.value);
        slideNumber(SlideState.number);
        callSlideParts();
        form.reset();
    });
}

function setupLegacyButton() {
    const el = document.getElementById("go-legacy");
    el.addEventListener("click", () => {
        window.location.href = "legacy.html";
    });
}