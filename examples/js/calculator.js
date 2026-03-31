// Calculator — WinForms-style GUI built from JavaScript
let form = gui.createForm("Calculator");
gui.setProperty(form, "Width", 280);
gui.setProperty(form, "Height", 400);

// Display
gui.addControl(form, "TextBox", "display", 10, 10, 250, 40);
gui.setProperty("display", "Text", "0");

// State
let current = "0";
let previous = "";
let operator = "";
let resetNext = false;

function updateDisplay() {
    gui.setProperty("display", "Text", current);
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

// Helper: create a button handler that captures the value at creation time
function makeDigitHandler(d) {
    return () => { pressDigit(d); };
}

function makeOpHandler(op) {
    return () => { pressOperator(op); };
}

// Button layout
let buttons = ["7", "8", "9", "/", "4", "5", "6", "*", "1", "2", "3", "-", "C", "0", "=", "+"];

let row = 0;
let col = 0;
for (let i = 0; i < 16; i++) {
    let label = buttons[i];
    let btnName = "btn" + i;
    let x = 10 + col * 63;
    let y = 60 + row * 55;
    gui.addControl(form, "Button", btnName, x, y, 58, 48);
    gui.setProperty(btnName, "Text", label);

    if (label === "C") {
        gui.onEvent(btnName, "Click", () => { pressClear(); });
    } else if (label === "=") {
        gui.onEvent(btnName, "Click", () => { calculate(); });
    } else if (label === "+" || label === "-" || label === "*" || label === "/") {
        gui.onEvent(btnName, "Click", makeOpHandler(label));
    } else {
        gui.onEvent(btnName, "Click", makeDigitHandler(label));
    }

    col = col + 1;
    if (col >= 4) { col = 0; row = row + 1; }
}

gui.runApplication(form);
