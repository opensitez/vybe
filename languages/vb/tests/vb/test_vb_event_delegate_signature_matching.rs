use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Event & Delegate Signature Matching & Overloads
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_delegate_custom_signature_sub() {
    let src = r#"
Delegate Sub MathOp(a As Integer, b As Integer, ByRef result As Integer)

Module Program
    Private Sub Add(a As Integer, b As Integer, ByRef result As Integer)
        result = a + b
    End Sub

    Sub Main()
        Dim op As MathOp = AddressOf Add
        Dim res As Integer = 0
        op(10, 20, res)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_delegate_custom_signature_function() {
    let src = r#"
Delegate Function Transform(input As String) As String

Module Program
    Private Function Upper(s As String) As String
        Return s.ToUpper()
    End Function

    Sub Main()
        Dim t As Transform = AddressOf Upper
        Console.WriteLine(t("hello"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["HELLO"]);
}

#[test]
fn test_vb_delegate_multicast_invocation_list() {
    let src = r#"
Imports System

Delegate Sub MultiNotify(msg As String)

Module Program
    Private Sub Logger1(msg As String) : Console.WriteLine("Log1: " & msg) : End Sub
    Private Sub Logger2(msg As String) : Console.WriteLine("Log2: " & msg) : End Sub

    Sub Main()
        Dim d As MultiNotify = AddressOf Logger1
        d = CType([Delegate].Combine(d, New MultiNotify(AddressOf Logger2)), MultiNotify)
        d("Message")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Log1: Message", "Log2: Message"]);
}

#[test]
fn test_vb_delegate_function_multicast_returns_last_result() {
    let src = r#"
Imports System

Delegate Function Compute(x As Integer) As Integer

Module Program
    Private Function Square(x As Integer) As Integer : Return x * x : End Function
    Private Function Cube(x As Integer) As Integer : Return x * x * x : End Function

    Sub Main()
        Dim c As Compute = AddressOf Square
        c = CType([Delegate].Combine(c, New Compute(AddressOf Cube)), Compute)
        Console.WriteLine(c(3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["27"]);
}

#[test]
fn test_vb_delegate_generic_t_result() {
    let src = r#"
Delegate Function Mapper(Of TIn, TOut)(item As TIn) As TOut

Module Program
    Sub Main()
        Dim m As Mapper(Of Integer, String) = Function(n) "Number_" & n
        Console.WriteLine(m(42))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Number_42"]);
}

#[test]
fn test_vb_delegate_covariance_return_type() {
    let src = r#"
Imports System

Class Animal : End Class
Class Dog : Inherits Animal : End Class

Delegate Function AnimalFactory() As Animal

Module Program
    Private Function CreateDog() As Dog
        Return New Dog()
    End Function

    Sub Main()
        Dim f As AnimalFactory = AddressOf CreateDog
        Dim a As Animal = f()
        Console.WriteLine(a IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_delegate_contravariance_parameter_type() {
    let src = r#"
Imports System

Class Animal : End Class
Class Dog : Inherits Animal : End Class

Delegate Sub DogHandler(d As Dog)

Module Program
    Private Sub ProcessAnimal(a As Animal)
        Console.WriteLine("Processed Animal: " & a.GetType().Name)
    End Sub

    Sub Main()
        Dim h As DogHandler = AddressOf ProcessAnimal
        h(New Dog())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Processed Animal: Dog"]);
}

#[test]
fn test_vb_event_matching_custom_delegate() {
    let src = r#"
Delegate Sub StatusChangeHandler(oldStatus As String, newStatus As String)

Class Machine
    Public Event StatusChanged As StatusChangeHandler
    Public Sub UpdateStatus(n As String)
        RaiseEvent StatusChanged("Offline", n)
    End Sub
End Class

Module Program
    Sub Main()
        Dim m As New Machine()
        AddHandler m.StatusChanged, Sub(o, n) Console.WriteLine(o & " -> " & n)
        m.UpdateStatus("Online")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Offline -> Online"]);
}

#[test]
fn test_vb_delegate_begin_invoke_end_invoke_async_pattern() {
    let src = r#"
Imports System

Delegate Function SlowCalc(n As Integer) As Integer

Module Program
    Private Function Compute(n As Integer) As Integer
        Return n * 10
    End Function

    Sub Main()
        Dim calc As SlowCalc = AddressOf Compute
        ' Invoke synchronously to test signature
        Dim res = calc.Invoke(5)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_delegate_dynamic_invoke() {
    let src = r#"
Imports System

Delegate Function MathFunc(a As Double, b As Double) As Double

Module Program
    Private Function Add(a As Double, b As Double) As Double
        Return a + b
    End Function

    Sub Main()
        Dim mf As [Delegate] = New MathFunc(AddressOf Add)
        Dim res = mf.DynamicInvoke(12.5, 7.5)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_delegate_equality_operator() {
    let src = r#"
Imports System

Delegate Sub SimpleAction()

Module Program
    Private Sub Action1() : End Sub
    Private Sub Action2() : End Sub

    Sub Main()
        Dim d1 As SimpleAction = AddressOf Action1
        Dim d2 As SimpleAction = AddressOf Action1
        Dim d3 As SimpleAction = AddressOf Action2
        Console.WriteLine((d1 = d2) & "|" & (d1 = d3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_delegate_target_property_instance_binding() {
    let src = r#"
Imports System

Class TargetService
    Public Prefix As String
    Public Sub PrintMsg(msg As String)
        Console.WriteLine(Prefix & ": " & msg)
    End Sub
End Class

Delegate Sub MsgDelegate(msg As String)

Module Program
    Sub Main()
        Dim ts As New TargetService With {.Prefix = "LOG"}
        Dim d As MsgDelegate = AddressOf ts.PrintMsg
        Console.WriteLine(d.Target Is ts)
        d("Message text")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "LOG: Message text"]);
}

#[test]
fn test_vb_delegate_method_info_inspection() {
    let src = r#"
Imports System
Imports System.Reflection

Delegate Function StringOp(input As String) As Integer

Module Program
    Private Function GetLen(s As String) As Integer
        Return s.Length
    End Function

    Sub Main()
        Dim op As StringOp = AddressOf GetLen
        Dim mi As MethodInfo = op.Method
        Console.WriteLine(mi.Name & "|" & mi.ReturnType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["GetLen|Int32"]);
}

#[test]
fn test_vb_event_handler_relaxation_allowing_fewer_parameters() {
    let src = r#"
Imports System

Class EventPublisher
    Public Event Click As EventHandler
    Public Sub Fire()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    ' VB.NET allows delegate relaxation: omitting parameters
    Private Sub ParameterlessHandler()
        Console.WriteLine("Parameterless Handler Invoked")
    End Sub

    Sub Main()
        Dim ep As New EventPublisher()
        AddHandler ep.Click, AddressOf ParameterlessHandler
        ep.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Parameterless Handler Invoked"]);
}

#[test]
fn test_vb_delegate_relaxation_function_to_sub() {
    let src = r#"
Imports System

Class Button
    Public Event Click As EventHandler
    Public Sub Fire()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    ' VB.NET allows binding a Function returning a value to a Sub delegate, ignoring the return value
    Private Function FuncHandler(sender As Object, e As EventArgs) As Integer
        Console.WriteLine("FuncHandler Invoked")
        Return 42
    End Function

    Sub Main()
        Dim b As New Button()
        AddHandler b.Click, AddressOf FuncHandler
        b.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FuncHandler Invoked"]);
}

#[test]
fn test_vb_delegate_nested_delegate_type_in_interface() {
    let src = r#"
Interface IProcessor
    Delegate Sub ResultHandler(success As Boolean, payload As String)
    Sub Process(callback As ResultHandler)
End Interface

Class ConcreteProcessor
    Implements IProcessor
    Public Sub Process(callback As IProcessor.ResultHandler) Implements IProcessor.Process
        callback(True, "OK")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IProcessor = New ConcreteProcessor()
        p.Process(Sub(s, msg) Console.WriteLine(s & ":" & msg))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:OK"]);
}

#[test]
fn test_vb_delegate_struct_method_target() {
    let src = r#"
Imports System

Structure Calculator
    Public Factor As Integer
    Public Sub New(f As Integer)
        Factor = f
    End Sub
    Public Function Multiply(val As Integer) As Integer
        Return val * Factor
    End Function
End Structure

Delegate Function MultiplyDel(val As Integer) As Integer

Module Program
    Sub Main()
        Dim c As New Calculator(5)
        Dim d As MultiplyDel = AddressOf c.Multiply
        Console.WriteLine(d(10))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_delegate_combining_null_delegate() {
    let src = r#"
Imports System

Delegate Sub SimpleDel()

Module Program
    Private Sub Target() : Console.WriteLine("Target") : End Sub

    Sub Main()
        Dim d1 As SimpleDel = Nothing
        Dim d2 As SimpleDel = AddressOf Target
        Dim combined As SimpleDel = CType([Delegate].Combine(d1, d2), SimpleDel)
        combined()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Target"]);
}

#[test]
fn test_vb_delegate_removing_all_leaves_null() {
    let src = r#"
Imports System

Delegate Sub SimpleDel()

Module Program
    Private Sub Target() : End Sub

    Sub Main()
        Dim d As SimpleDel = AddressOf Target
        d = CType([Delegate].Remove(d, AddressOf Target), SimpleDel)
        Console.WriteLine(d Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_delegate_conversion_to_action() {
    let src = r#"
Imports System

Module Program
    Private Sub Output(msg As String)
        Console.WriteLine("ActionMsg: " & msg)
    End Sub

    Sub Main()
        Dim act As Action(Of String) = AddressOf Output
        act("Hello Action")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ActionMsg: Hello Action"]);
}
