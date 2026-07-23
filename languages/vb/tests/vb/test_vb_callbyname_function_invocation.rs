use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Microsoft.VisualBasic.Interaction.CallByName Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_callbyname_method_invocation_no_args() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Worker
        Public Function GetStatus() As String
            Return "Active"
        End Function
    End Class

    Sub Main()
        Dim w As New Worker()
        Dim res = CallByName(w, "GetStatus", CallType.Method)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Active"]);
}

#[test]
fn test_vb_callbyname_method_invocation_with_args() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Calculator
        Public Function Multiply(a As Integer, b As Integer) As Integer
            Return a * b
        End Function
    End Class

    Sub Main()
        Dim calc As New Calculator()
        Dim res = CallByName(calc, "Multiply", CallType.Method, 6, 7)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_callbyname_property_get() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class User
        Public Property Username As String = "Alice"
    End Class

    Sub Main()
        Dim u As New User()
        Dim name = CallByName(u, "Username", CallType.Get)
        Console.WriteLine(name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_callbyname_property_let_set() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class User
        Public Property Username As String
    End Class

    Sub Main()
        Dim u As New User()
        CallByName(u, "Username", CallType.Set, "Bob")
        Console.WriteLine(u.Username)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bob"]);
}

#[test]
fn test_vb_callbyname_sub_procedure_call() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class ActionRunner
        Public Sub ExecuteTask(taskName As String)
            Console.WriteLine("Task Executed: " & taskName)
        End Sub
    End Class

    Sub Main()
        Dim runner As New ActionRunner()
        CallByName(runner, "ExecuteTask", CallType.Method, "Backup")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Task Executed: Backup"]);
}

#[test]
fn test_vb_callbyname_case_insensitive_name_matching() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Target
        Public Function SampleMethod() As String
            Return "MatchFound"
        End Function
    End Class

    Sub Main()
        Dim t As New Target()
        ' CallByName in VB.NET is case-insensitive for member names!
        Dim res = CallByName(t, "samplemethod", CallType.Method)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["MatchFound"]);
}

#[test]
fn test_vb_callbyname_inherited_member_call() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Parent
        Public Function ParentFunc() As String
            Return "ParentResult"
        End Function
    End Class

    Class Child
        Inherits Parent
    End Class

    Sub Main()
        Dim c As New Child()
        Dim res = CallByName(c, "ParentFunc", CallType.Method)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ParentResult"]);
}

#[test]
fn test_vb_callbyname_overridden_virtual_member() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class BaseDevice
        Public Overridable Function GetDeviceType() As String
            Return "Base"
        End Function
    End Class

    Class PrinterDevice
        Inherits BaseDevice
        Public Overrides Function GetDeviceType() As String
            Return "Printer"
        End Function
    End Class

    Sub Main()
        Dim dev As BaseDevice = New PrinterDevice()
        Dim res = CallByName(dev, "GetDeviceType", CallType.Method)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Printer"]);
}

#[test]
fn test_vb_callbyname_missing_member_throws_exception() {
    let src = r#"
Imports System
Imports Microsoft.VisualBasic

Module Program
    Class EmptyClass
    End Class

    Sub Main()
        Dim obj As New EmptyClass()
        Try
            CallByName(obj, "NonExistent", CallType.Method)
        Catch ex As Exception
            Console.WriteLine("CallByName Missing Member Exception Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["CallByName Missing Member Exception Caught"]
    );
}

#[test]
fn test_vb_callbyname_null_target_throws_exception() {
    let src = r#"
Imports System
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Try
            CallByName(Nothing, "AnyMethod", CallType.Method)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null CallByName Instance")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null CallByName Instance"]
    );
}

#[test]
fn test_vb_callbyname_property_let_value_type() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Counter
        Public Property Value As Integer
    End Class

    Sub Main()
        Dim c As New Counter()
        CallByName(c, "Value", CallType.Let, 100)
        Console.WriteLine(c.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_callbyname_indexed_property_get() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Catalog
        Private items As String() = {"Alpha", "Beta", "Gamma"}
        Default Public Property Item(idx As Integer) As String
            Get
                Return items(idx)
            End Get
        End Property
    End Class

    Sub Main()
        Dim cat As New Catalog()
        Dim res = CallByName(cat, "Item", CallType.Get, 1)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Beta"]);
}

#[test]
fn test_vb_callbyname_return_array_result() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Provider
        Public Function GetTags() As String()
            Return New String() {"tag1", "tag2"}
        End Function
    End Class

    Sub Main()
        Dim p As New Provider()
        Dim tags As String() = CType(CallByName(p, "GetTags", CallType.Method), String())
        Console.WriteLine(String.Join(",", tags))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["tag1,tag2"]);
}

#[test]
fn test_vb_callbyname_overloaded_method_resolution() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Formatter
        Public Function Format(val As Integer) As String
            Return "Int:" & val
        End Function
        Public Function Format(val As String) As String
            Return "Str:" & val
        End Function
    End Class

    Sub Main()
        Dim f As New Formatter()
        Dim res1 = CallByName(f, "Format", CallType.Method, 99)
        Dim res2 = CallByName(f, "Format", CallType.Method, "Text")
        Console.WriteLine(res1 & "|" & res2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int:99|Str:Text"]);
}

#[test]
fn test_vb_callbyname_byref_argument_mutation() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Transformer
        Public Sub DoubleValue(ByRef x As Integer)
            x *= 2
        End Sub
    End Class

    Sub Main()
        Dim t As New Transformer()
        Dim val As Integer = 25
        CallByName(t, "DoubleValue", CallType.Method, val)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_callbyname_optional_arguments_omitted() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Class Config
        Public Function Build(host As String, Optional port As Integer = 80) As String
            Return host & ":" & port
        End Function
    End Class

    Sub Main()
        Dim c As New Config()
        Dim res = CallByName(c, "Build", CallType.Method, "localhost")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["localhost:80"]);
}

#[test]
fn test_vb_callbyname_custom_structure_instance() {
    let src = r#"
Imports Microsoft.VisualBasic

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Function GetSum() As Integer
        Return X + Y
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Point With {.X = 10, .Y = 20}
        Dim res = CallByName(p, "GetSum", CallType.Method)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_callbyname_interface_implementation_call() {
    let src = r#"
Imports Microsoft.VisualBasic

Interface ICommand
    Function Run() As Boolean
End Interface

Class ConcreteCommand
    Implements ICommand
    Public Function Run() As Boolean Implements ICommand.Run
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim cmd As ICommand = New ConcreteCommand()
        Dim res = CallByName(cmd, "Run", CallType.Method)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_callbyname_property_returns_custom_class() {
    let src = r#"
Imports Microsoft.VisualBasic

Class SubSystem
    Public Property Name As String = "Core"
End Class

Class RootSystem
    Public Property SubSys As New SubSystem()
End Class

Module Program
    Sub Main()
        Dim r As New RootSystem()
        Dim subObj = CallByName(r, "SubSys", CallType.Get)
        Dim subName = CallByName(subObj, "Name", CallType.Get)
        Console.WriteLine(subName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Core"]);
}

#[test]
fn test_vb_callbyname_method_raising_exception_unwrapped() {
    let src = r#"
Imports System
Imports Microsoft.VisualBasic

Class FaultyComponent
    Public Sub Crash()
        Throw New InvalidOperationException("Component Crashed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim comp As New FaultyComponent()
        Try
            CallByName(comp, "Crash", CallType.Method)
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Component Crashed"]);
}
