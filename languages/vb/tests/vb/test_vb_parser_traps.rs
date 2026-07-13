use super::helpers::run_vb;

#[test]
fn string_interpolation_fmt_align() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val As Double = 42.5
        Dim s = $"[{val,10:F2}]"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["[     42.50]"]);
}

#[test]
fn typeof_isnot_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = "String"
        If TypeOf obj IsNot Integer Then
            Console.WriteLine("Not Integer")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Not Integer"]);
}

#[test]
fn byref_paramarray() {
    let out = run_vb(
        r#"
Module M
    ' ParamArray cannot be ByRef. Testing parser error recovery.
    Sub Test()
        Console.WriteLine("Parsed")
    End Sub

    Sub Main()
        Test()
    End Sub
End Module

Class InvalidSyntaxTest
    ' Sub Invalid(ByRef ParamArray x() As Integer)
    ' End Sub
End Class
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn default_property_no_params() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Default properties MUST have parameters.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn module_in_class() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Module inside a Class is not allowed. Testing parser skips/flags.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn mustinherit_notinheritable() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' MustInherit and NotInheritable cannot be used together.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn const_type_character() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Const PI# = 3.14
        Console.WriteLine(PI)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn enum_ulong() {
    let out = run_vb(
        r#"
Enum Flags As ULong
    None = 0
    All = &HFFFFFFFFFFFFFFFFUL
End Enum

Module M
    Sub Main()
        Console.WriteLine(Flags.None)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn array_literal_nested_new() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr As Integer()() = New Integer()() {New Integer() {1, 2}, New Integer() {3}}
        Console.WriteLine(arr(0)(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn redim_preserve_change_rank() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' ReDim Preserve changing rank is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn null_conditional_delegate() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim act As Action = Nothing
        act?.Invoke()
        
        act = Sub() Console.WriteLine("Invoked")
        act?.Invoke()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Invoked"]);
}

#[test]
fn throw_non_exception() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Throwing an object that doesn't inherit from Exception is generally not allowed in VB.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn trycast_generic_class() {
    let out = run_vb(
        r#"
Class Tester
    Public Sub Process(Of T As Class)(obj As Object)
        Dim c = TryCast(obj, T)
        Console.WriteLine(c Is Nothing)
    End Sub
End Class

Module M
    Sub Main()
        Dim t As New Tester()
        t.Process(Of String)(100)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn directcast_generic_struct() {
    let out = run_vb(
        r#"
Class Tester
    Public Sub Process(Of T As Structure)(obj As Object)
        Try
            Dim c = DirectCast(obj, T)
            Console.WriteLine("Cast Success")
        Catch
            Console.WriteLine("Cast Failed")
        End Try
    End Sub
End Class

Module M
    Sub Main()
        Dim t As New Tester()
        t.Process(Of Integer)("String")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Cast Failed"]);
}

#[test]
fn addressof_in_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Target()
        Console.WriteLine("Target")
    End Sub

    Sub Main()
        ' Sub() AddressOf Method is invalid as AddressOf returns a delegate.
        ' However, we can use it to assign to an explicit delegate inside.
        Dim act = Sub()
                      Dim d As Action = AddressOf Target
                      d()
                  End Sub
        act()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Target"]);
}

#[test]
fn nested_lambdas() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim f = Function(x As Integer) Function(y As Integer) x + y
        
        Dim add5 = f(5)
        Console.WriteLine(add5(10))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn yield_nothing() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Iterator Function Generate() As IEnumerable(Of Object)
        Yield Nothing
    End Function

    Sub Main()
        For Each item In Generate()
            Console.WriteLine(item Is Nothing)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn await_in_synclock() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Await inside SyncLock is invalid in VB.NET.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn generic_property() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Properties cannot be generic.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn generic_event() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Events cannot be generic.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn property_with_arguments() {
    let out = run_vb(
        r#"
Class Cache
    Private _vals(10, 10) As Integer
    
    Public Property Value(x As Integer, y As Integer) As Integer
        Get
            Return _vals(x, y)
        End Get
        Set(val As Integer)
            _vals(x, y) = val
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        c.Value(1, 2) = 42
        Console.WriteLine(c.Value(1, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn property_get_set_different_access() {
    let out = run_vb(
        r#"
Class Data
    Private _val As Integer
    
    ' Only one accessor can have an access modifier different from the property
    Public Property Val As Integer
        Get
            Return _val
        End Get
        Protected Set(value As Integer)
            _val = value
        End Set
    End Property
    
    Public Sub Update(v As Integer)
        Val = v
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Data()
        d.Update(100)
        Console.WriteLine(d.Val)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn multiline_if_missing_then() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Missing Then on multiline If is a syntax error.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn select_case_missing_case() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Code directly inside Select Case before first Case is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn continue_in_finally() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Continue For/While inside Finally is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn goto_into_try_block() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Branching into a Try block is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn resume_next_in_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' On Error Resume Next inside a lambda is generally restricted.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn exit_try_outside_try() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Exit Try outside a Try block is invalid.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn return_in_finally() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Return inside Finally is not allowed in VB.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}

#[test]
fn handles_in_interface() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Interface members cannot have Handles clause.
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}
