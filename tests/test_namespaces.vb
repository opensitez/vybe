' Test System.Console
System.Console.WriteLine("Hello from System.Console")

' Test System.Math
Dim maxVal As Double
maxVal = System.Math.Max(10, 20)
System.Console.WriteLine("Max(10, 20) = " & maxVal)

Dim sqrtVal As Double
sqrtVal = System.Math.Sqrt(16)
System.Console.WriteLine("Sqrt(16) = " & sqrtVal)

' Test Object Assignment
Dim myMath As Object
myMath = System.Math
Dim minVal As Double
minVal = myMath.Min(10, 20)
System.Console.WriteLine("myMath.Min(10, 20) = " & minVal)

Dim myConsole As Object
myConsole = System.Console
myConsole.WriteLine("Hello from myConsole")
