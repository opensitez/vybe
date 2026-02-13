' Simple Hello World vybe Program

Sub btnHello_Click()
    MsgBox("Hello, World!")
End Sub

Sub btnCalculate_Click()
    Dim x As Integer
    Dim y As Integer
    Dim result As Integer

    x = 10
    y = 20
    result = x + y

    MsgBox("Result: " & result)
End Sub

Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function

Sub TestFunction()
    Dim sum As Integer
    sum = Add(5, 3)
    MsgBox("5 + 3 = " & sum)
End Sub
