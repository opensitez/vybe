use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Reflection MethodInfo MakeGenericMethod & Dynamic Invocation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_reflection_make_generic_method_single_type_arg() {
    let src = r#"
Class GenericCalculator
    Public Function Identity(Of T)(val As T) As T
        Return val
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New GenericCalculator()
        Dim openMethod = GetType(GenericCalculator).GetMethod("Identity")
        Dim closedMethod = openMethod.MakeGenericMethod(GetType(Integer))
        Dim res = closedMethod.Invoke(calc, {42})
        Console.WriteLine(res.GetType().Name & "=" & res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32=42"]);
}

#[test]
fn test_vb_reflection_make_generic_method_two_type_args() {
    let src = r#"
Class Mapper
    Public Function Map(Of T1, T2)(item As T1, transform As System.Func(Of T1, T2)) As T2
        Return transform(item)
    End Function
End Class

Module Program
    Sub Main()
        Dim m As New Mapper()
        Dim openMethod = GetType(Mapper).GetMethod("Map")
        Dim closedMethod = openMethod.MakeGenericMethod(GetType(String), GetType(Integer))
        Dim fn As System.Func(Of String, Integer) = Function(s) s.Length
        Dim res = closedMethod.Invoke(m, {"VisualBasic", fn})
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11"]);
}

#[test]
fn test_vb_reflection_is_generic_method_definition() {
    let src = r#"
Class Utility
    Public Function Process(Of T)(item As T) As String : Return item.ToString() : End Function
    Public Function NonGeneric(item As String) As String : Return item : End Function
End Class

Module Program
    Sub Main()
        Dim mGen = GetType(Utility).GetMethod("Process")
        Dim mNonGen = GetType(Utility).GetMethod("NonGeneric")
        Console.WriteLine(mGen.IsGenericMethodDefinition & "|" & mNonGen.IsGenericMethodDefinition)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_reflection_get_generic_arguments_method_info() {
    let src = r#"
Class Service
    Public Sub Execute(Of T1, T2)() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Service).GetMethod("Execute")
        Dim typeArgs = m.GetGenericArguments()
        Console.WriteLine(typeArgs.Length & ":" & typeArgs(0).Name & "," & typeArgs(1).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:T1,T2"]);
}

#[test]
fn test_vb_reflection_shared_static_generic_method_invoke() {
    let src = r#"
Class Helper
    Public Shared Function Wrap(Of T)(item As T) As String
        Return "[" & item.ToString() & "]"
    End Function
End Class

Module Program
    Sub Main()
        Dim m = GetType(Helper).GetMethod("Wrap").MakeGenericMethod(GetType(Double))
        Dim res = m.Invoke(Nothing, {3.14})
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[3.14]"]);
}

#[test]
fn test_vb_reflection_invoke_method_with_byref_parameter() {
    let src = r#"
Class Processor
    Public Sub DoubleValue(ByRef num As Integer)
        num *= 2
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Processor()
        Dim m = GetType(Processor).GetMethod("DoubleValue")
        Dim args As Object() = {25}
        m.Invoke(p, args)
        Console.WriteLine(args(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_reflection_invoke_method_with_optional_parameters() {
    let src = r#"
Imports System.Reflection

Class Printer
    Public Function PrintMsg(msg As String, Optional prefix As String = "LOG:") As String
        Return prefix & " " & msg
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Printer()
        Dim m = GetType(Printer).GetMethod("PrintMsg")
        Dim res = m.Invoke(p, BindingFlags.OptionalParamBinding, Nothing, {"Hello", Missing.Value}, Nothing)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LOG: Hello"]);
}

#[test]
fn test_vb_reflection_invoke_method_with_paramarray() {
    let src = r#"
Class Aggregator
    Public Function SumAll(ParamArray numbers As Integer()) As Integer
        Dim sum = 0
        For Each n In numbers : sum += n : Next
        Return sum
    End Function
End Class

Module Program
    Sub Main()
        Dim agg As New Aggregator()
        Dim m = GetType(Aggregator).GetMethod("SumAll")
        Dim res = m.Invoke(agg, {New Integer() {10, 20, 30}})
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_reflection_private_method_invocation_binding_flags() {
    let src = r#"
Imports System.Reflection

Class InternalProcessor
    Private Function SecretFormula(x As Integer) As Integer
        Return x * 7
    End Function
End Class

Module Program
    Sub Main()
        Dim proc As New InternalProcessor()
        Dim m = GetType(InternalProcessor).GetMethod("SecretFormula", BindingFlags.Instance Or BindingFlags.NonPublic)
        Dim res = m.Invoke(proc, {5})
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["35"]);
}

#[test]
fn test_vb_reflection_overloaded_methods_resolution() {
    let src = r#"
Class OverloadSample
    Public Function Compute(x As Integer) As String : Return "Int_" & x : End Function
    Public Function Compute(x As String) As String : Return "Str_" & x : End Function
End Class

Module Program
    Sub Main()
        Dim s As New OverloadSample()
        Dim mInt = GetType(OverloadSample).GetMethod("Compute", {GetType(Integer)})
        Dim mStr = GetType(OverloadSample).GetMethod("Compute", {GetType(String)})

        Console.WriteLine(mInt.Invoke(s, {10}) & "|" & mStr.Invoke(s, {"abc"}))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int_10|Str_abc"]);
}

#[test]
fn test_vb_reflection_method_info_return_type() {
    let src = r#"
Class Sample
    Public Function GetName() As String : Return "" : End Function
    Public Sub DoNothing() : End Sub
End Class

Module Program
    Sub Main()
        Dim m1 = GetType(Sample).GetMethod("GetName")
        Dim m2 = GetType(Sample).GetMethod("DoNothing")
        Console.WriteLine(m1.ReturnType.Name & "|" & (m2.ReturnType Is GetType(Void)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["String|True"]);
}

#[test]
fn test_vb_reflection_method_info_is_virtual_is_abstract() {
    let src = r#"
MustInherit Class BaseClass
    Public MustOverride Sub AbstractMethod()
    Public Overridable Sub VirtualMethod() : End Sub
End Class

Module Program
    Sub Main()
        Dim mAbs = GetType(BaseClass).GetMethod("AbstractMethod")
        Dim mVirt = GetType(BaseClass).GetMethod("VirtualMethod")
        Console.WriteLine(mAbs.IsAbstract & "|" & mVirt.IsVirtual)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_reflection_method_info_create_delegate() {
    let src = r#"
Imports System

Class ActionRunner
    Public Function Execute(msg As String) As String
        Return "Executed: " & msg
    End Function
End Class

Module Program
    Sub Main()
        Dim runner As New ActionRunner()
        Dim m = GetType(ActionRunner).GetMethod("Execute")
        Dim del = CType(m.CreateDelegate(GetType(Func(Of String, String)), runner), Func(Of String, String))
        Console.WriteLine(del("DirectDelegate"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Executed: DirectDelegate"]);
}

#[test]
fn test_vb_reflection_method_info_generic_constraint_check() {
    let src = r#"
Class ConstrainedMethod
    Public Sub Process(Of T As Class)(item As T) : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(ConstrainedMethod).GetMethod("Process")
        Dim tParam = m.GetGenericArguments()(0)
        Dim constraints = tParam.GetGenericParameterConstraints()
        Console.WriteLine(tParam.GenericParameterAttributes.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ReferenceTypeConstraint"]);
}

#[test]
fn test_vb_reflection_method_info_custom_attribute() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Method)>
Class RouteAttribute
    Inherits Attribute
    Public Path As String
    Public Sub New(p As String) : Path = p : End Sub
End Class

Class Controller
    <Route("/api/users")>
    Public Sub GetUsers() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Controller).GetMethod("GetUsers")
        Dim attr = CType(m.GetCustomAttributes(GetType(RouteAttribute), False)(0), RouteAttribute)
        Console.WriteLine(attr.Path)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["/api/users"]);
}

#[test]
fn test_vb_reflection_method_info_extension_method_check() {
    let src = r#"
Imports System.Reflection
Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function ReverseString(s As String) As String
        Dim chars = s.ToCharArray()
        Array.Reverse(chars)
        Return New String(chars)
    End Function
End Module

Module Program
    Sub Main()
        Dim m = GetType(StringExtensions).GetMethod("ReverseString")
        Dim isExt = m.IsDefined(GetType(ExtensionAttribute), False)
        Console.WriteLine(isExt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_reflection_method_info_invoke_throws_target_invocation_exception() {
    let src = r#"
Imports System
Imports System.Reflection

Class FaultyMethod
    Public Sub Fail()
        Throw New ArgumentException("Invalid Argument")
    End Sub
End Class

Module Program
    Sub Main()
        Dim fm As New FaultyMethod()
        Dim m = GetType(FaultyMethod).GetMethod("Fail")
        Try
            m.Invoke(fm, Nothing)
        Catch ex As TargetInvocationException
            Console.WriteLine(ex.InnerException.GetType().Name & ": " & ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentException: Invalid Argument"]);
}

#[test]
fn test_vb_reflection_method_info_tuple_return_value() {
    let src = r#"
Class TupleService
    Public Function GetPair() As (Code As Integer, Name As String)
        Return (200, "OK")
    End Function
End Class

Module Program
    Sub Main()
        Dim svc As New TupleService()
        Dim m = GetType(TupleService).GetMethod("GetPair")
        Dim res As (Integer, String) = CType(m.Invoke(svc, Nothing), (Integer, String))
        Console.WriteLine(res.Item1 & " " & res.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200 OK"]);
}

#[test]
fn test_vb_reflection_method_info_anonymous_type_return() {
    let src = r#"
Class AnonService
    Public Function GetAnon() As Object
        Return New With {.Status = "AnonSuccess"}
    End Function
End Class

Module Program
    Sub Main()
        Dim svc As New AnonService()
        Dim m = GetType(AnonService).GetMethod("GetAnon")
        Dim res As Dynamic = m.Invoke(svc, Nothing)
        Console.WriteLine(res.Status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AnonSuccess"]);
}

#[test]
fn test_vb_reflection_method_info_get_base_definition() {
    let src = r#"
Class BaseClass
    Public Overridable Sub Display() : End Sub
End Class

Class DerivedClass
    Inherits BaseClass
    Public Overrides Sub Display() : End Sub
End Class

Module Program
    Sub Main()
        Dim mDerived = GetType(DerivedClass).GetMethod("Display")
        Dim mBase = mDerived.GetBaseDefinition()
        Console.WriteLine(mBase.DeclaringType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BaseClass"]);
}
