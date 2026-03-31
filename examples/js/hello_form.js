// Hello Form — a WinForms-style GUI with event handling, created from JavaScript
let form = gui.createForm("Hello from JS!");
gui.setProperty(form, "Width", 450);
gui.setProperty(form, "Height", 350);

// Add a label
gui.addControl(form, "Label", "label1", 20, 20, 400, 30);
gui.setProperty("label1", "Text", "Hello World from JavaScript!");

// Add a counter label
gui.addControl(form, "Label", "counterLabel", 20, 60, 400, 30);
gui.setProperty("counterLabel", "Text", "Click count: 0");

// Add a button with click handler
gui.addControl(form, "Button", "btn1", 20, 100, 150, 35);
gui.setProperty("btn1", "Text", "Click Me!");

let clickCount = 0;
gui.onEvent("btn1", "Click", () => {
    clickCount = clickCount + 1;
    gui.setProperty("counterLabel", "Text", "Click count: " + clickCount);
    gui.setProperty("btn1", "Text", "Clicked " + clickCount + "x");
});

// Add a reset button
gui.addControl(form, "Button", "btnReset", 180, 100, 150, 35);
gui.setProperty("btnReset", "Text", "Reset");

gui.onEvent("btnReset", "Click", () => {
    clickCount = 0;
    gui.setProperty("counterLabel", "Text", "Click count: 0");
    gui.setProperty("btn1", "Text", "Click Me!");
});

// Add a textbox
gui.addControl(form, "TextBox", "txt1", 20, 150, 300, 25);
gui.setProperty("txt1", "Text", "Type here...");

// Info label
gui.addControl(form, "Label", "info", 20, 200, 400, 50);
gui.setProperty("info", "Text", "This form was created entirely from JavaScript, running on the Vybe bytecode VM with WinForms rendering!");

// Launch!
console.log("Launching JS form with event handling...");
gui.runApplication(form);
