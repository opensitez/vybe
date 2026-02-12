///! Tests for class fields, methods, constructors, multiple instances, and integration.

use irys_runtime::{Interpreter, RuntimeSideEffect};
use irys_parser::ast::Identifier;
use irys_parser::parse_program;

fn run_vb(code: &str) -> Vec<String> {
    let program = parse_program(code).expect("Parse error");
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");
    interp.call_procedure(&Identifier::new("Main"), &[]).expect("Failed to call Main");
    interp.side_effects.iter().filter_map(|e| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = e { Some(msg.trim_end().to_string()) } else { None }
    }).collect()
}

#[test]
fn test_class_fields_and_methods() {
    let output = run_vb(r#"
Public Class Rectangle
    Public Width As Double
    Public Height As Double

    Public Function Area() As Double
        Return Width * Height
    End Function

    Public Function Perimeter() As Double
        Return 2 * (Width + Height)
    End Function
End Class

Sub Main()
    Dim r As New Rectangle()
    r.Width = 5.0
    r.Height = 3.0
    Console.WriteLine("Area: " & r.Area())
    Console.WriteLine("Perimeter: " & r.Perimeter())
End Sub
"#);
    assert_eq!(output, vec!["Area: 15", "Perimeter: 16"]);
}

#[test]
fn test_class_constructor_default_values() {
    let output = run_vb(r#"
Public Class Counter
    Public Value As Integer = 0

    Public Sub Increment()
        Value = Value + 1
    End Sub

    Public Sub IncrementBy(amount As Integer)
        Value = Value + amount
    End Sub
End Class

Sub Main()
    Dim c As New Counter()
    Console.WriteLine(c.Value)
    c.Increment()
    c.Increment()
    c.IncrementBy(10)
    Console.WriteLine(c.Value)
End Sub
"#);
    assert_eq!(output, vec!["0", "12"]);
}

#[test]
fn test_multiple_class_instances() {
    let output = run_vb(r#"
Public Class Point
    Public X As Integer
    Public Y As Integer
End Class

Sub Main()
    Dim p1 As New Point()
    p1.X = 1
    p1.Y = 2
    Dim p2 As New Point()
    p2.X = 10
    p2.Y = 20
    Console.WriteLine(p1.X & "," & p1.Y)
    Console.WriteLine(p2.X & "," & p2.Y)
    ' Ensure they're independent
    p1.X = 100
    Console.WriteLine(p1.X)
    Console.WriteLine(p2.X)
End Sub
"#);
    assert_eq!(output, vec!["1,2", "10,20", "100", "10"]);
}

#[test]
fn test_class_integration() {
    let output = run_vb(r#"
Public Class Accumulator
    Public Total As Double = 0
    Public CallCount As Integer = 0

    Public Sub Add(value As Double)
        Total = Total + value
        CallCount = CallCount + 1
    End Sub

    Public Function Average() As Double
        If CallCount = 0 Then
            Return 0
        End If
        Return Total / CallCount
    End Function

    Public Function Describe() As String
        Return "Total=" & Total & " Count=" & CallCount
    End Function
End Class

Sub Main()
    Dim acc As New Accumulator()
    Console.WriteLine(acc.Describe())
    acc.Add(10)
    acc.Add(20)
    acc.Add(30)
    Console.WriteLine(acc.Describe())
    Console.WriteLine("Avg: " & acc.Average())
End Sub
"#);
    assert_eq!(output[0], "Total=0 Count=0");
    assert_eq!(output[1], "Total=60 Count=3");
    assert_eq!(output[2], "Avg: 20");
}
