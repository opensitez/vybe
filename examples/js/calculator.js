// Calculator — built against the WHATWG DOM, which is what JavaScript has.
// A control IS an element: `createElement` makes it, `appendChild` puts it in
// the page, `addEventListener` binds the click. `document` is a global the web
// platform binds, exactly as `window.document` is in a browser.
document.setTitle("Calculator");

// **The page styles itself.** Nothing below computes a coordinate or a size:
// the rules say what the boxes are and the engine lays them out, which is the
// whole point of going through the DOM rather than a widget factory. Swap the
// engine for a browser and this is unchanged.
let sheet = document.createElement("style");
sheet.appendChild(document.createTextNode(
    "body { background: #202020; }" +
    "#display { width: 252px; height: 56px; margin: 8px; font-size: 32px; }" +
    ".row { margin-left: 8px; }" +
    ".key { width: 60px; height: 56px; margin: 2px; font-size: 20px; }"
));
document.body.appendChild(sheet);

let display = document.createElement("input");
display.setAttribute("id", "display");
display.setAttribute("value", "0");
document.body.appendChild(display);

// State
let current = "0";
let previous = "";
let operator = "";
let resetNext = false;

function updateDisplay() {
    display.setAttribute("value", current);
}

function pressDigit(d) {
    if (resetNext) {
        current = d;
        resetNext = false;
    } else {
        if (current === "0") {
            current = d;
        } else {
            current = current + d;
        }
    }
    updateDisplay();
}

function pressOperator(op) {
    if (previous !== "" && !resetNext) {
        calculate();
    }
    previous = current;
    operator = op;
    resetNext = true;
}

function calculate() {
    if (previous === "" || operator === "") { return; }
    let a = parseFloat(previous);
    let b = parseFloat(current);
    let result = 0;
    if (operator === "+") { result = a + b; }
    if (operator === "-") { result = a - b; }
    if (operator === "*") { result = a * b; }
    if (operator === "/") {
        if (b === 0) {
            current = "Error";
            previous = "";
            operator = "";
            resetNext = true;
            updateDisplay();
            return;
        }
        result = a / b;
    }
    current = "" + result;
    previous = "";
    operator = "";
    resetNext = true;
    updateDisplay();
}

function pressClear() {
    current = "0";
    previous = "";
    operator = "";
    resetNext = false;
    updateDisplay();
}

// Each handler captures its own value at creation time.
function makeDigitHandler(d) {
    return () => { pressDigit(d); };
}

function makeOpHandler(op) {
    return () => { pressOperator(op); };
}

let keys = ["7", "8", "9", "/", "4", "5", "6", "*", "1", "2", "3", "-", "C", "0", "=", "+"];

// Four keys per row. A row is a `<div>`; the buttons inside it are inline, so
// the browser lays the grid out — nothing here computes an x or a y.
let row = null;
for (let i = 0; i < 16; i++) {
    if (i % 4 === 0) {
        row = document.createElement("div");
        row.setAttribute("class", "row");
        document.body.appendChild(row);
    }

    let label = keys[i];
    let key = document.createElement("button");
    key.setAttribute("class", "key");
    key.appendChild(document.createTextNode(label));
    row.appendChild(key);

    if (label === "C") {
        key.addEventListener("click", () => { pressClear(); });
    } else if (label === "=") {
        key.addEventListener("click", () => { calculate(); });
    } else if (label === "+" || label === "-" || label === "*" || label === "/") {
        key.addEventListener("click", makeOpHandler(label));
    } else {
        key.addEventListener("click", makeDigitHandler(label));
    }
}
