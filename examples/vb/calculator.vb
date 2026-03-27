' VB Calculator with WinForms-style GUI
' Run: vybec examples/vb/calculator.vb

Module Program
    Dim displayName As String = ""
    Dim currentText As String = "0"
    Dim previousValue As String = ""
    Dim currentOp As String = ""
    Dim resetNext As Boolean = False

    Sub UpdateDisplay()
        vybe.gui.setProperty(displayName, "Text", currentText)
    End Sub

    Sub PressDigit(d As String)
        If resetNext Then
            currentText = d
            resetNext = False
        Else
            If currentText = "0" Then
                currentText = d
            Else
                currentText = currentText & d
            End If
        End If
        UpdateDisplay()
    End Sub

    Sub PressOperator(op As String)
        If previousValue <> "" And Not resetNext Then
            DoCalculate()
        End If
        previousValue = currentText
        currentOp = op
        resetNext = True
    End Sub

    Sub DoCalculate()
        If previousValue = "" Or currentOp = "" Then Return
        Dim a As Double = Val(previousValue)
        Dim b As Double = Val(currentText)
        Dim result As Double = 0
        If currentOp = "+" Then result = a + b
        If currentOp = "-" Then result = a - b
        If currentOp = "*" Then result = a * b
        If currentOp = "/" Then
            If b = 0 Then
                currentText = "Error"
                previousValue = ""
                currentOp = ""
                resetNext = True
                UpdateDisplay()
                Return
            End If
            result = a / b
        End If
        currentText = CStr(result)
        previousValue = ""
        currentOp = ""
        resetNext = True
        UpdateDisplay()
    End Sub

    Sub PressClear()
        currentText = "0"
        previousValue = ""
        currentOp = ""
        resetNext = False
        UpdateDisplay()
    End Sub

    Sub OnBtn7()
        PressDigit("7")
    End Sub
    Sub OnBtn8()
        PressDigit("8")
    End Sub
    Sub OnBtn9()
        PressDigit("9")
    End Sub
    Sub OnBtnDiv()
        PressOperator("/")
    End Sub
    Sub OnBtn4()
        PressDigit("4")
    End Sub
    Sub OnBtn5()
        PressDigit("5")
    End Sub
    Sub OnBtn6()
        PressDigit("6")
    End Sub
    Sub OnBtnMul()
        PressOperator("*")
    End Sub
    Sub OnBtn1()
        PressDigit("1")
    End Sub
    Sub OnBtn2()
        PressDigit("2")
    End Sub
    Sub OnBtn3()
        PressDigit("3")
    End Sub
    Sub OnBtnSub()
        PressOperator("-")
    End Sub
    Sub OnBtnC()
        PressClear()
    End Sub
    Sub OnBtn0()
        PressDigit("0")
    End Sub
    Sub OnBtnEq()
        DoCalculate()
    End Sub
    Sub OnBtnAdd()
        PressOperator("+")
    End Sub

    Sub Main()
        Dim form As Object = Window.Forms.Form("Calculator")

        ' Display
        Dim display As Object = Window.Forms.TextBox()
        display.text = "0"
        display.left = 10
        display.top = 10
        display.width = 250
        display.height = 40
        display.readonly = True
        form.Controls.Add(display)
        displayName = display.name

        ' Row 1: 7 8 9 /
        Dim btn7 As Object = Window.Forms.Button()
        btn7.text = "7"
        btn7.left = 10
        btn7.top = 60
        btn7.width = 58
        btn7.height = 48
        form.Controls.Add(btn7)
        AddHandler btn7.Click, AddressOf OnBtn7

        Dim btn8 As Object = Window.Forms.Button()
        btn8.text = "8"
        btn8.left = 73
        btn8.top = 60
        btn8.width = 58
        btn8.height = 48
        form.Controls.Add(btn8)
        AddHandler btn8.Click, AddressOf OnBtn8

        Dim btn9 As Object = Window.Forms.Button()
        btn9.text = "9"
        btn9.left = 136
        btn9.top = 60
        btn9.width = 58
        btn9.height = 48
        form.Controls.Add(btn9)
        AddHandler btn9.Click, AddressOf OnBtn9

        Dim btnDiv As Object = Window.Forms.Button()
        btnDiv.text = "/"
        btnDiv.left = 199
        btnDiv.top = 60
        btnDiv.width = 58
        btnDiv.height = 48
        form.Controls.Add(btnDiv)
        AddHandler btnDiv.Click, AddressOf OnBtnDiv

        ' Row 2: 4 5 6 *
        Dim btn4 As Object = Window.Forms.Button()
        btn4.text = "4"
        btn4.left = 10
        btn4.top = 115
        btn4.width = 58
        btn4.height = 48
        form.Controls.Add(btn4)
        AddHandler btn4.Click, AddressOf OnBtn4

        Dim btn5 As Object = Window.Forms.Button()
        btn5.text = "5"
        btn5.left = 73
        btn5.top = 115
        btn5.width = 58
        btn5.height = 48
        form.Controls.Add(btn5)
        AddHandler btn5.Click, AddressOf OnBtn5

        Dim btn6 As Object = Window.Forms.Button()
        btn6.text = "6"
        btn6.left = 136
        btn6.top = 115
        btn6.width = 58
        btn6.height = 48
        form.Controls.Add(btn6)
        AddHandler btn6.Click, AddressOf OnBtn6

        Dim btnMul As Object = Window.Forms.Button()
        btnMul.text = "*"
        btnMul.left = 199
        btnMul.top = 115
        btnMul.width = 58
        btnMul.height = 48
        form.Controls.Add(btnMul)
        AddHandler btnMul.Click, AddressOf OnBtnMul

        ' Row 3: 1 2 3 -
        Dim btn1 As Object = Window.Forms.Button()
        btn1.text = "1"
        btn1.left = 10
        btn1.top = 170
        btn1.width = 58
        btn1.height = 48
        form.Controls.Add(btn1)
        AddHandler btn1.Click, AddressOf OnBtn1

        Dim btn2 As Object = Window.Forms.Button()
        btn2.text = "2"
        btn2.left = 73
        btn2.top = 170
        btn2.width = 58
        btn2.height = 48
        form.Controls.Add(btn2)
        AddHandler btn2.Click, AddressOf OnBtn2

        Dim btn3 As Object = Window.Forms.Button()
        btn3.text = "3"
        btn3.left = 136
        btn3.top = 170
        btn3.width = 58
        btn3.height = 48
        form.Controls.Add(btn3)
        AddHandler btn3.Click, AddressOf OnBtn3

        Dim btnSubtr As Object = Window.Forms.Button()
        btnSubtr.text = "-"
        btnSubtr.left = 199
        btnSubtr.top = 170
        btnSubtr.width = 58
        btnSubtr.height = 48
        form.Controls.Add(btnSubtr)
        AddHandler btnSubtr.Click, AddressOf OnBtnSub

        ' Row 4: C 0 = +
        Dim btnC As Object = Window.Forms.Button()
        btnC.text = "C"
        btnC.left = 10
        btnC.top = 225
        btnC.width = 58
        btnC.height = 48
        form.Controls.Add(btnC)
        AddHandler btnC.Click, AddressOf OnBtnC

        Dim btn0 As Object = Window.Forms.Button()
        btn0.text = "0"
        btn0.left = 73
        btn0.top = 225
        btn0.width = 58
        btn0.height = 48
        form.Controls.Add(btn0)
        AddHandler btn0.Click, AddressOf OnBtn0

        Dim btnEq As Object = Window.Forms.Button()
        btnEq.text = "="
        btnEq.left = 136
        btnEq.top = 225
        btnEq.width = 58
        btnEq.height = 48
        form.Controls.Add(btnEq)
        AddHandler btnEq.Click, AddressOf OnBtnEq

        Dim btnAdd As Object = Window.Forms.Button()
        btnAdd.text = "+"
        btnAdd.left = 199
        btnAdd.top = 225
        btnAdd.width = 58
        btnAdd.height = 48
        form.Controls.Add(btnAdd)
        AddHandler btnAdd.Click, AddressOf OnBtnAdd

        Application.Run(form)
    End Sub
End Module
