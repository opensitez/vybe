' VB Calculator with WinForms GUI
' Run: vybec examples/vb/calculator.vb

Imports System.Windows.Forms

Module Program
    Dim display As TextBox
    Dim currentText As String = "0"
    Dim previousValue As String = ""
    Dim currentOp As String = ""
    Dim resetNext As Boolean = False

    Sub UpdateDisplay()
        display.Text = currentText
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
        Dim form As New Form()

        ' Display
        display = New TextBox()
        display.Text = "0"
        display.Left = 10
        display.Top = 10
        display.Width = 260
        display.Height = 40
        display.ReadOnly = True
        form.Controls.Add(display)

        ' Row 1: 7 8 9 /
        Dim btn7 As New Button()
        btn7.Text = "7"
        btn7.Left = 10
        btn7.Top = 60
        btn7.Width = 58
        btn7.Height = 48
        form.Controls.Add(btn7)
        AddHandler btn7.Click, AddressOf OnBtn7

        Dim btn8 As New Button()
        btn8.Text = "8"
        btn8.Left = 73
        btn8.Top = 60
        btn8.Width = 58
        btn8.Height = 48
        form.Controls.Add(btn8)
        AddHandler btn8.Click, AddressOf OnBtn8

        Dim btn9 As New Button()
        btn9.Text = "9"
        btn9.Left = 136
        btn9.Top = 60
        btn9.Width = 58
        btn9.Height = 48
        form.Controls.Add(btn9)
        AddHandler btn9.Click, AddressOf OnBtn9

        Dim btnDiv As New Button()
        btnDiv.Text = "/"
        btnDiv.Left = 199
        btnDiv.Top = 60
        btnDiv.Width = 58
        btnDiv.Height = 48
        form.Controls.Add(btnDiv)
        AddHandler btnDiv.Click, AddressOf OnBtnDiv

        ' Row 2: 4 5 6 *
        Dim btn4 As New Button()
        btn4.Text = "4"
        btn4.Left = 10
        btn4.Top = 115
        btn4.Width = 58
        btn4.Height = 48
        form.Controls.Add(btn4)
        AddHandler btn4.Click, AddressOf OnBtn4

        Dim btn5 As New Button()
        btn5.Text = "5"
        btn5.Left = 73
        btn5.Top = 115
        btn5.Width = 58
        btn5.Height = 48
        form.Controls.Add(btn5)
        AddHandler btn5.Click, AddressOf OnBtn5

        Dim btn6 As New Button()
        btn6.Text = "6"
        btn6.Left = 136
        btn6.Top = 115
        btn6.Width = 58
        btn6.Height = 48
        form.Controls.Add(btn6)
        AddHandler btn6.Click, AddressOf OnBtn6

        Dim btnMul As New Button()
        btnMul.Text = "*"
        btnMul.Left = 199
        btnMul.Top = 115
        btnMul.Width = 58
        btnMul.Height = 48
        form.Controls.Add(btnMul)
        AddHandler btnMul.Click, AddressOf OnBtnMul

        ' Row 3: 1 2 3 -
        Dim btn1 As New Button()
        btn1.Text = "1"
        btn1.Left = 10
        btn1.Top = 170
        btn1.Width = 58
        btn1.Height = 48
        form.Controls.Add(btn1)
        AddHandler btn1.Click, AddressOf OnBtn1

        Dim btn2 As New Button()
        btn2.Text = "2"
        btn2.Left = 73
        btn2.Top = 170
        btn2.Width = 58
        btn2.Height = 48
        form.Controls.Add(btn2)
        AddHandler btn2.Click, AddressOf OnBtn2

        Dim btn3 As New Button()
        btn3.Text = "3"
        btn3.Left = 136
        btn3.Top = 170
        btn3.Width = 58
        btn3.Height = 48
        form.Controls.Add(btn3)
        AddHandler btn3.Click, AddressOf OnBtn3

        Dim btnSubtr As New Button()
        btnSubtr.Text = "-"
        btnSubtr.Left = 199
        btnSubtr.Top = 170
        btnSubtr.Width = 58
        btnSubtr.Height = 48
        form.Controls.Add(btnSubtr)
        AddHandler btnSubtr.Click, AddressOf OnBtnSub

        ' Row 4: C 0 = +
        Dim btnC As New Button()
        btnC.Text = "C"
        btnC.Left = 10
        btnC.Top = 225
        btnC.Width = 58
        btnC.Height = 48
        form.Controls.Add(btnC)
        AddHandler btnC.Click, AddressOf OnBtnC

        Dim btn0 As New Button()
        btn0.Text = "0"
        btn0.Left = 73
        btn0.Top = 225
        btn0.Width = 58
        btn0.Height = 48
        form.Controls.Add(btn0)
        AddHandler btn0.Click, AddressOf OnBtn0

        Dim btnEq As New Button()
        btnEq.Text = "="
        btnEq.Left = 136
        btnEq.Top = 225
        btnEq.Width = 58
        btnEq.Height = 48
        form.Controls.Add(btnEq)
        AddHandler btnEq.Click, AddressOf OnBtnEq

        Dim btnAdd As New Button()
        btnAdd.Text = "+"
        btnAdd.Left = 199
        btnAdd.Top = 225
        btnAdd.Width = 58
        btnAdd.Height = 48
        form.Controls.Add(btnAdd)
        AddHandler btnAdd.Click, AddressOf OnBtnAdd

        Application.Run(form)
    End Sub
End Module
