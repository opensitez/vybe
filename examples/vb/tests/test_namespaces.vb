' Test System.Console
System.Console.WriteLine("Hello from System.Console")

' Test System.Math
Dim maxVal As Double
maxVal = System.Math.Max(10, 20)
System.Console.WriteLine("Max(10, 20) = " & maxVal)

Dim sqrtVal As Double
sqrtVal = System.Math.Sqrt(16)
System.Console.WriteLine("Sqrt(16) = " & sqrtVal)

' Test Object Assignment — late binding through Object.
'
' The receiver is an INSTANCE of a namespaced type, not the type itself.
' `myMath = System.Math` is not VB: a class has no value, and the compiler
' says so — BC30109, "'Math' is a class type and cannot be used as an
' expression" (verified against dotnet 10). This half of the file used to
' assign `System.Math` and `System.Console` and assert on the result.
Dim myBuilder As Object
myBuilder = New System.Text.StringBuilder()
myBuilder.Append("Hello from ")
myBuilder.Append("System.Text.StringBuilder")
System.Console.WriteLine(myBuilder.ToString())

Dim myList As Object
myList = New System.Collections.ArrayList()
myList.Add("first")
myList.Add("second")
System.Console.WriteLine("ArrayList Count = " & myList.Count)
System.Console.WriteLine("ArrayList(0) = " & myList(0))
