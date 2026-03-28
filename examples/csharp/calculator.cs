// C# Calculator with WinForms-style GUI
// Run: vybec examples/csharp/calculator.cs

using System;
using System.Windows.Forms;
using System.Drawing;

public class Calculator : Form
{
    private TextBox txtDisplay;
    private string currentText = "0";
    private string previousValue = "";
    private string currentOp = "";
    private bool resetNext = false;

    public Calculator()
    {
        this.Text = "C# Calculator";
        this.ClientSize = new Size(280, 320);

        // Display
        txtDisplay = new TextBox();
        txtDisplay.Name = "txtDisplay";
        txtDisplay.Text = "0";
        txtDisplay.Location = new Point(10, 10);
        txtDisplay.Size = new Size(260, 40);
        txtDisplay.ReadOnly = true;
        this.Controls.Add(txtDisplay);

        // Row 1: 7 8 9 /
        AddButton("btn7", "7", 10, 60, 60, 48);
        AddButton("btn8", "8", 73, 60, 60, 48);
        AddButton("btn9", "9", 136, 60, 60, 48);
        AddButton("btnDiv", "/", 199, 60, 60, 48);

        // Row 2: 4 5 6 *
        AddButton("btn4", "4", 10, 115, 60, 48);
        AddButton("btn5", "5", 73, 115, 60, 48);
        AddButton("btn6", "6", 136, 115, 60, 48);
        AddButton("btnMul", "*", 199, 115, 60, 48);

        // Row 3: 1 2 3 -
        AddButton("btn1", "1", 10, 170, 60, 48);
        AddButton("btn2", "2", 73, 170, 60, 48);
        AddButton("btn3", "3", 136, 170, 60, 48);
        AddButton("btnSub", "-", 199, 170, 60, 48);

        // Row 4: C 0 = +
        AddButton("btnC", "C", 10, 225, 60, 48);
        AddButton("btn0", "0", 73, 225, 60, 48);
        AddButton("btnEq", "=", 136, 225, 60, 48);
        AddButton("btnAdd", "+", 199, 225, 60, 48);
    }

    private void AddButton(string name, string text, int x, int y, int w, int h)
    {
        var btn = new Button();
        btn.Name = name;
        btn.Text = text;
        btn.Location = new Point(x, y);
        btn.Size = new Size(w, h);
        this.Controls.Add(btn);
        btn.Click += this.HandleClick;
    }

    private void HandleClick(object sender, object e)
    {
        // Dispatch by button name
        var name = sender;
        OnButtonClick("" + name);
    }

    private void UpdateDisplay()
    {
        txtDisplay.Text = currentText;
    }

    private void PressDigit(string d)
    {
        if (resetNext)
        {
            currentText = d;
            resetNext = false;
        }
        else
        {
            if (currentText == "0")
                currentText = d;
            else
                currentText = currentText + d;
        }
        UpdateDisplay();
    }

    private void PressOperator(string op)
    {
        if (previousValue != "" && !resetNext)
            DoCalculate();
        previousValue = currentText;
        currentOp = op;
        resetNext = true;
    }

    private void DoCalculate()
    {
        if (previousValue == "" || currentOp == "") return;

        double a = Convert.ToDouble(previousValue);
        double b = Convert.ToDouble(currentText);
        double result = 0;

        if (currentOp == "+") result = a + b;
        if (currentOp == "-") result = a - b;
        if (currentOp == "*") result = a * b;
        if (currentOp == "/")
        {
            if (b == 0)
            {
                currentText = "Error";
                previousValue = "";
                currentOp = "";
                resetNext = true;
                UpdateDisplay();
                return;
            }
            result = a / b;
        }

        currentText = "" + result;
        previousValue = "";
        currentOp = "";
        resetNext = true;
        UpdateDisplay();
    }

    private void PressClear()
    {
        currentText = "0";
        previousValue = "";
        currentOp = "";
        resetNext = false;
        UpdateDisplay();
    }

    // Event handler dispatch — called by the form runner
    public void OnButtonClick(string buttonName)
    {
        switch (buttonName)
        {
            case "btn0": PressDigit("0"); break;
            case "btn1": PressDigit("1"); break;
            case "btn2": PressDigit("2"); break;
            case "btn3": PressDigit("3"); break;
            case "btn4": PressDigit("4"); break;
            case "btn5": PressDigit("5"); break;
            case "btn6": PressDigit("6"); break;
            case "btn7": PressDigit("7"); break;
            case "btn8": PressDigit("8"); break;
            case "btn9": PressDigit("9"); break;
            case "btnAdd": PressOperator("+"); break;
            case "btnSub": PressOperator("-"); break;
            case "btnMul": PressOperator("*"); break;
            case "btnDiv": PressOperator("/"); break;
            case "btnEq": DoCalculate(); break;
            case "btnC": PressClear(); break;
        }
    }
}

// Entry point
var calc = new Calculator();
Application.Run(calc);
