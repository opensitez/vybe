use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Object Late-Bound Method Calls & Member Dispatch
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_late_bound_method_call_single_arg() {
    let src = r#"
Module Program
    Class Calculator
        Public Function AddTen(x As Integer) As Integer
            Return x + 10
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Calculator()
        Dim res As Integer = CInt(obj.AddTen(5))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_late_bound_method_call_multiple_args() {
    let src = r#"
Module Program
    Class TextFormatter
        Public Function ConcatStrings(a As String, b As String, c As String) As String
            Return a & "-" & b & "-" & c
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New TextFormatter()
        Dim res As String = CStr(obj.ConcatStrings("A", "B", "C"))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A-B-C"]);
}

#[test]
fn test_vb_late_bound_sub_invocation() {
    let src = r#"
Module Program
    Class Logger
        Public Sub LogMessage(msg As String)
            Console.WriteLine("LOG: " & msg)
        End Sub
    End Class

    Sub Main()
        Dim obj As Object = New Logger()
        obj.LogMessage("System Startup")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LOG: System Startup"]);
}

#[test]
fn test_vb_late_bound_method_overload_resolution() {
    let src = r#"
Module Program
    Class Printer
        Public Function Print(x As Integer) As String
            Return "Int:" & x
        End Function

        Public Function Print(x As String) As String
            Return "Str:" & x
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Printer()
        Console.WriteLine(CStr(obj.Print(42)) & "|" & CStr(obj.Print("Hello")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int:42|Str:Hello"]);
}

#[test]
fn test_vb_late_bound_method_optional_parameters() {
    let src = r#"
Module Program
    Class Greeter
        Public Function Greet(name As String, Optional prefix As String = "Hello") As String
            Return prefix & " " & name
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Greeter()
        Console.WriteLine(CStr(obj.Greet("Alice")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello Alice"]);
}

#[test]
fn test_vb_late_bound_method_paramarray_args() {
    let src = r#"
Module Program
    Class MathOps
        Public Function SumAll(ParamArray numbers As Integer()) As Integer
            Dim sum = 0
            For Each n In numbers
                sum += n
            Next
            Return sum
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New MathOps()
        Dim total As Integer = CInt(obj.SumAll(10, 20, 30))
        Console.WriteLine(total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_late_bound_method_return_array() {
    let src = r#"
Module Program
    Class Provider
        Public Function GetNumbers() As Integer()
            Return New Integer() {1, 2, 3}
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Provider()
        Dim arr As Integer() = CType(obj.GetNumbers(), Integer())
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_late_bound_method_byref_parameter() {
    let src = r#"
Module Program
    Class Mutator
        Public Sub Increment(ByRef val As Integer)
            val += 5
        End Sub
    End Class

    Sub Main()
        Dim obj As Object = New Mutator()
        Dim num As Integer = 10
        obj.Increment(num)
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_late_bound_inherited_method_call() {
    let src = r#"
Module Program
    Class BaseClass
        Public Function BaseMethod() As String
            Return "BaseMethodResult"
        End Function
    End Class

    Class DerivedClass
        Inherits BaseClass
    End Class

    Sub Main()
        Dim obj As Object = New DerivedClass()
        Console.WriteLine(CStr(obj.BaseMethod()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BaseMethodResult"]);
}

#[test]
fn test_vb_late_bound_overridden_virtual_method() {
    let src = r#"
Module Program
    Class Animal
        Public Overridable Function Speak() As String
            Return "Generic Sound"
        End Function
    End Class

    Class Dog
        Inherits Animal
        Public Overrides Function Speak() As String
            Return "Bark"
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Dog()
        Console.WriteLine(CStr(obj.Speak()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bark"]);
}

#[test]
fn test_vb_late_bound_method_missing_member_exception() {
    let src = r#"
Imports System

Module Program
    Class Dummy
    End Class

    Sub Main()
        Dim obj As Object = New Dummy()
        Try
            obj.NonExistentMethod()
        Catch ex As Exception
            Console.WriteLine("Missing Member Exception Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Missing Member Exception Caught"]);
}

#[test]
fn test_vb_late_bound_method_null_reference_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Try
            obj.SomeMethod()
        Catch ex As NullReferenceException
            Console.WriteLine("NullReferenceException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NullReferenceException Caught"]);
}

#[test]
fn test_vb_late_bound_interface_implementation_call() {
    let src = r#"
Module Program
    Interface IService
        Function Execute() As String
    End Interface

    Class ServiceImpl
        Implements IService
        Public Function Execute() As String Implements IService.Execute
            Return "Executed"
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New ServiceImpl()
        Console.WriteLine(CStr(obj.Execute()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Executed"]);
}

#[test]
fn test_vb_late_bound_generic_method_invocation() {
    let src = r#"
Module Program
    Class GenericWorker
        Public Function Identity(Of T)(val As T) As T
            Return val
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New GenericWorker()
        Dim res As String = CStr(obj.Identity("GenericVal"))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["GenericVal"]);
}

#[test]
fn test_vb_late_bound_structure_method_call() {
    let src = r#"
Module Program
    Structure Vector2D
        Public X As Integer
        Public Y As Integer
        Public Function GetMagnitude() As Double
            Return Math.Sqrt(X * X + Y * Y)
        End Function
    End Structure

    Sub Main()
        Dim obj As Object = New Vector2D With {.X = 3, .Y = 4}
        Console.WriteLine(CDbl(obj.GetMagnitude()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_late_bound_chained_object_calls() {
    let src = r#"
Module Program
    Class Node
        Public Property NextNode As Node
        Public Function GetValue() As String
            Return "NodeValue"
        End Function
    End Class

    Sub Main()
        Dim first As New Node With {.NextNode = New Node()}
        Dim obj As Object = first
        Console.WriteLine(CStr(obj.NextNode.GetValue()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NodeValue"]);
}

#[test]
fn test_vb_late_bound_call_with_named_arguments() {
    let src = r#"
Module Program
    Class Configurator
        Public Function Format(prefix As String, suffix As String) As String
            Return prefix & ":" & suffix
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Configurator()
        Dim res = CStr(obj.Format(suffix:="END", prefix:="START"))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["START:END"]);
}

#[test]
fn test_vb_late_bound_method_raising_exception() {
    let src = r#"
Imports System

Module Program
    Class Faulty
        Public Sub Fail()
            Throw New InvalidOperationException("Custom Fault")
        End Sub
    End Class

    Sub Main()
        Dim obj As Object = New Faulty()
        Try
            obj.Fail()
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Custom Fault"]);
}

#[test]
fn test_vb_late_bound_value_type_conversion_args() {
    let src = r#"
Module Program
    Class ImplicitConverter
        Public Function DoubleVal(x As Double) As Double
            Return x * 2.0
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New ImplicitConverter()
        ' Integer 10 should implicitly convert to Double 10.0 in late binding!
        Dim res As Double = CDbl(obj.DoubleVal(10))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_late_bound_to_string_override() {
    let src = r#"
Module Program
    Class Person
        Public Property Name As String
        Public Overrides Function ToString() As String
            Return "Person:" & Name
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New Person With {.Name = "Bob"}
        Console.WriteLine(obj.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Person:Bob"]);
}
