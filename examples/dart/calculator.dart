// Calculator — WinForms-style GUI built from Dart
void main() {
    var form = gui.createForm("Calculator");
    gui.setProperty(form, "Width", 280);
    gui.setProperty(form, "Height", 400);

    // Display
    gui.addControl(form, "TextBox", "display", 10, 10, 250, 40);
    gui.setProperty("display", "Text", "0");

    // State
    var current = "0";
    var previous = "";
    var operator = "";
    var resetNext = false;

    void updateDisplay() {
        gui.setProperty("display", "Text", current);
    }

    void pressDigit(String d) {
        if (resetNext) {
            current = d;
            resetNext = false;
        } else {
            if (current == "0") {
                current = d;
            } else {
                current = current + d;
            }
        }
        updateDisplay();
    }

    void calculate() {
        if (previous == "" || operator == "") { return; }
        var a = previous - 0;
        var b = current - 0;
        var result = 0;
        if (operator == "+") { result = a + b; }
        if (operator == "-") { result = a - b; }
        if (operator == "*") { result = a * b; }
        if (operator == "/") {
            if (b == 0) {
                current = "Error";
                previous = "";
                operator = "";
                resetNext = true;
                updateDisplay();
                return;
            }
            result = a / b;
        }
        current = "$result";
        previous = "";
        operator = "";
        resetNext = true;
        updateDisplay();
    }

    void pressOperator(String op) {
        if (previous != "" && !resetNext) {
            calculate();
        }
        previous = current;
        operator = op;
        resetNext = true;
    }

    void pressClear() {
        current = "0";
        previous = "";
        operator = "";
        resetNext = false;
        updateDisplay();
    }

    Function makeDigitHandler(String d) {
        return () { pressDigit(d); };
    }

    Function makeOpHandler(String op) {
        return () { pressOperator(op); };
    }

    // Button layout
    var buttons = ["7", "8", "9", "/", "4", "5", "6", "*", "1", "2", "3", "-", "C", "0", "=", "+"];

    var row = 0;
    var col = 0;
    for (var i = 0; i < 16; i++) {
        var label = buttons[i];
        var btnName = "btn$i";
        var x = 10 + col * 63;
        var y = 60 + row * 55;
        gui.addControl(form, "Button", btnName, x, y, 58, 48);
        gui.setProperty(btnName, "Text", label);

        if (label == "C") {
            gui.onEvent(btnName, "Click", () { pressClear(); });
        } else if (label == "=") {
            gui.onEvent(btnName, "Click", () { calculate(); });
        } else if (label == "+" || label == "-" || label == "*" || label == "/") {
            gui.onEvent(btnName, "Click", makeOpHandler(label));
        } else {
            gui.onEvent(btnName, "Click", makeDigitHandler(label));
        }

        col = col + 1;
        if (col >= 4) { col = 0; row = row + 1; }
    }

    gui.runApplication(form);
}
