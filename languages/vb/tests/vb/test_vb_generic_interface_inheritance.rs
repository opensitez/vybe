use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Interface Inheritance & Multi-Parameter Constraints
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_interface_inheritance_specialization() {
    let src = r#"
Interface IReadRepository(Of T)
    Function GetById(id As Integer) As T
End Interface

Interface IWriteRepository(Of T)
    Sub Save(entity As T)
End Interface

Interface IRepository(Of T)
    Inherits IReadRepository(Of T), IWriteRepository(Of T)
End Interface

Class UserRepo
    Implements IRepository(Of String)
    Private user As String = ""
    Public Function GetById(id As Integer) As String Implements IReadRepository(Of String).GetById
        Return user
    End Function
    Public Sub Save(entity As String) Implements IWriteRepository(Of String).Save
        user = entity
    End Sub
End Class

Module Program
    Sub Main()
        Dim repo As IRepository(Of String) = New UserRepo()
        repo.Save("Alice")
        Console.WriteLine(repo.GetById(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_generic_interface_two_type_parameters() {
    let src = r#"
Interface IMapping(Of TKey, TValue)
    Function Map(key As TKey) As TValue
End Interface

Class IntToStringMapper
    Implements IMapping(Of Integer, String)
    Public Function Map(key As Integer) As String Implements IMapping(Of Integer, String).Map
        Return "Value_" & key
    End Function
End Class

Module Program
    Sub Main()
        Dim m As IMapping(Of Integer, String) = New IntToStringMapper()
        Console.WriteLine(m.Map(42))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value_42"]);
}

#[test]
fn test_vb_generic_interface_constraint_new_and_class() {
    let src = r#"
Interface IFactory(Of T As {Class, New})
    Function CreateInstance() As T
End Interface

Class ProductFactory(Of T As {Class, New})
    Implements IFactory(Of T)
    Public Function CreateInstance() As T Implements IFactory(Of T).CreateInstance
        Return New T()
    End Function
End Class

Class Car
    Public Model As String = "Sedan"
End Class

Module Program
    Sub Main()
        Dim f As IFactory(Of Car) = New ProductFactory(Of Car)()
        Dim c = f.CreateInstance()
        Console.WriteLine(c.Model)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Sedan"]);
}

#[test]
fn test_vb_generic_interface_constraint_structure_value_type() {
    let src = r#"
Interface INumberBox(Of T As Structure)
    Property Value As T
End Interface

Class IntBox
    Implements INumberBox(Of Integer)
    Public Property Value As Integer Implements INumberBox(Of Integer).Value
End Class

Module Program
    Sub Main()
        Dim b As INumberBox(Of Integer) = New IntBox() With {.Value = 99}
        Console.WriteLine(b.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_generic_interface_nested_generic_arguments() {
    let src = r#"
Imports System.Collections.Generic

Interface IBatchProcessor(Of TCollection As IEnumerable(Of String))
    Function ProcessBatch(batch As TCollection) As String
End Interface

Class ListBatchProcessor
    Implements IBatchProcessor(Of List(Of String))
    Public Function ProcessBatch(batch As List(Of String)) As String Implements IBatchProcessor(Of List(Of String)).ProcessBatch
        Return String.Join(",", batch)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IBatchProcessor(Of List(Of String)) = New ListBatchProcessor()
        Dim items As New List(Of String) From {"A", "B", "C"}
        Console.WriteLine(p.ProcessBatch(items))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B,C"]);
}

#[test]
fn test_vb_generic_interface_inheritance_with_additional_methods() {
    let src = r#"
Interface IService(Of T)
    Sub Execute(item As T)
End Interface

Interface IAdvancedService(Of T)
    Inherits IService(Of T)
    Sub ExecuteBatch(items As T())
End Interface

Class StringService
    Implements IAdvancedService(Of String)
    Public Sub Execute(item As String) Implements IService(Of String).Execute
        Console.WriteLine("Single: " & item)
    End Sub
    Public Sub ExecuteBatch(items As String()) Implements IAdvancedService(Of String).ExecuteBatch
        Console.WriteLine("Batch: " & String.Join("-", items))
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IAdvancedService(Of String) = New StringService()
        s.Execute("One")
        s.ExecuteBatch({"Two", "Three"})
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Single: One", "Batch: Two-Three"]);
}

#[test]
fn test_vb_generic_interface_covariant_out_return_type() {
    let src = r#"
Interface IProvider(Of Out T)
    Function GetItem() As T
End Interface

Class StringProvider
    Implements IProvider(Of String)
    Public Function GetItem() As String Implements IProvider(Of String).GetItem
        Return "Provided String"
    End Function
End Class

Module Program
    Sub Main()
        Dim provider As IProvider(Of Object) = New StringProvider()
        Console.WriteLine(provider.GetItem().ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Provided String"]);
}

#[test]
fn test_vb_generic_interface_contravariant_in_parameter() {
    let src = r#"
Interface IConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class ObjectConsumer
    Implements IConsumer(Of Object)
    Public Sub Consume(item As Object) Implements IConsumer(Of Object).Consume
        Console.WriteLine("Consuming: " & item.ToString())
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As IConsumer(Of String) = New ObjectConsumer()
        c.Consume("Test Message")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Consuming: Test Message"]);
}

#[test]
fn test_vb_generic_interface_multiple_interface_implementations_different_type_args() {
    let src = r#"
Interface IHandler(Of T)
    Sub Handle(item As T)
End Interface

Class DualHandler
    Implements IHandler(Of Integer), IHandler(Of String)
    Public Sub HandleInt(item As Integer) Implements IHandler(Of Integer).Handle
        Console.WriteLine("Int: " & item)
    End Sub
    Public Sub HandleString(item As String) Implements IHandler(Of String).Handle
        Console.WriteLine("String: " & item)
    End Sub
End Class

Module Program
    Sub Main()
        Dim dh As New DualHandler()
        Dim hInt As IHandler(Of Integer) = dh
        Dim hStr As IHandler(Of String) = dh
        hInt.Handle(10)
        hStr.Handle("Hello")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int: 10", "String: Hello"]);
}

#[test]
fn test_vb_generic_interface_self_referential_constraint() {
    let src = r#"
Imports System

Interface IComparableEntity(Of T As IComparableEntity(Of T))
    Function CompareTo(other As T) As Integer
End Interface

Class Account
    Implements IComparableEntity(Of Account)
    Public Property ID As Integer
    Public Function CompareTo(other As Account) As Integer Implements IComparableEntity(Of Account).CompareTo
        Return ID.CompareTo(other.ID)
    End Function
End Class

Module Program
    Sub Main()
        Dim a1 As New Account With {.ID = 10}
        Dim a2 As New Account With {.ID = 20}
        Console.WriteLine(a1.CompareTo(a2) < 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_generic_interface_explicit_implementation_shadowing() {
    let src = r#"
Interface IProcessor(Of T)
    Sub Process(data As T)
End Interface

Class Processor
    Implements IProcessor(Of String)
    Private Sub Process(data As String) Implements IProcessor(Of String).Process
        Console.WriteLine("Explicit Process: " & data)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IProcessor(Of String) = New Processor()
        p.Process("Data")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Explicit Process: Data"]);
}

#[test]
fn test_vb_generic_interface_method_overloads_with_generic_parameters() {
    let src = r#"
Interface ICalculator(Of T)
    Function Add(a As T, b As T) As T
End Interface

Class IntCalculator
    Implements ICalculator(Of Integer)
    Public Function Add(a As Integer, b As Integer) As Integer Implements ICalculator(Of Integer).Add
        Return a + b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As ICalculator(Of Integer) = New IntCalculator()
        Console.WriteLine(calc.Add(15, 25))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["40"]);
}

#[test]
fn test_vb_generic_interface_with_enum_type_parameter() {
    let src = r#"
Enum Mode
    Standard
    Advanced
End Enum

Interface IConfig(Of TEnum As Structure)
    Property CurrentMode As TEnum
End Interface

Class ModeConfig
    Implements IConfig(Of Mode)
    Public Property CurrentMode As Mode Implements IConfig(Of Mode).CurrentMode = Mode.Advanced
End Class

Module Program
    Sub Main()
        Dim cfg As IConfig(Of Mode) = New ModeConfig()
        Console.WriteLine(cfg.CurrentMode.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Advanced"]);
}

#[test]
fn test_vb_generic_interface_property_and_event() {
    let src = r#"
Imports System

Interface IObservableValue(Of T)
    Property Value As T
    Event ValueChanged As Action(Of T)
End Interface

Class ObservableInt
    Implements IObservableValue(Of Integer)
    Public Event ValueChanged As Action(Of Integer) Implements IObservableValue(Of Integer).ValueChanged
    Private _val As Integer
    Public Property Value As Integer Implements IObservableValue(Of Integer).Value
        Get
            Return _val
        Get
        End Get
        Set(val As Integer)
            _val = val
            RaiseEvent ValueChanged(_val)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim obs As IObservableValue(Of Integer) = New ObservableInt()
        AddHandler obs.ValueChanged, Sub(v) Console.WriteLine("New Value: " & v)
        obs.Value = 42
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["New Value: 42"]);
}

#[test]
fn test_vb_generic_interface_struct_implementer() {
    let src = r#"
Interface ISwap(Of T)
    Function SwapWith(other As T) As T
End Interface

Structure Pair(Of T)
    Implements ISwap(Of Pair(Of T))
    Public First As T
    Public Second As T
    Public Sub New(f As T, s As T)
        First = f : Second = s
    End Sub
    Public Function SwapWith(other As Pair(Of T)) As Pair(Of T) Implements ISwap(Of Pair(Of T)).SwapWith
        Return New Pair(Of T)(Second, First)
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of Integer)(10, 20)
        Dim swapped = p.SwapWith(p)
        Console.WriteLine(swapped.First & "," & swapped.Second)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,10"]);
}

#[test]
fn test_vb_generic_interface_indexer_property() {
    let src = r#"
Interface IGenericContainer(Of TKey, TValue)
    Default Property Item(key As TKey) As TValue
End Interface

Class SimpleMap
    Implements IGenericContainer(Of String, Integer)
    Default Public Property Item(key As String) As Integer Implements IGenericContainer(Of String, Integer).Item
        Get
            Return key.Length
        End Get
        Set(value As Integer)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim container As IGenericContainer(Of String, Integer) = New SimpleMap()
        Console.WriteLine(container("Hello"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_generic_interface_deep_type_argument_substitution() {
    let src = r#"
Imports System.Collections.Generic

Interface IDataPipeline(Of TIn, TOut)
    Function Process(input As IEnumerable(Of TIn)) As List(Of TOut)
End Interface

Class StringLengthPipeline
    Implements IDataPipeline(Of String, Integer)
    Public Function Process(input As IEnumerable(Of String)) As List(Of Integer) Implements IDataPipeline(Of String, Integer).Process
        Dim res As New List(Of Integer)()
        For Each s In input
            res.Add(s.Length)
        Next
        Return res
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IDataPipeline(Of String, Integer) = New StringLengthPipeline()
        Dim lengths = p.Process({"A", "BB", "CCC"})
        Console.WriteLine(String.Join(",", lengths))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_generic_interface_is_operator_check() {
    let src = r#"
Interface ICheckable(Of T)
End Interface

Class Impl
    Implements ICheckable(Of String)
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Impl()
        Console.WriteLine(TypeOf obj Is ICheckable(Of String))
        Console.WriteLine(TypeOf obj Is ICheckable(Of Integer))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_generic_interface_trycast_check() {
    let src = r#"
Interface IService(Of T)
    Sub Serve(t As T)
End Interface

Class ServiceImpl
    Implements IService(Of Integer)
    Public Sub Serve(t As Integer) Implements IService(Of Integer).Serve
        Console.WriteLine("Serving: " & t)
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New ServiceImpl()
        Dim s = TryCast(obj, IService(Of Integer))
        If s IsNot Nothing Then
            s.Serve(100)
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Serving: 100"]);
}

#[test]
fn test_vb_generic_interface_type_parameter_name_shadowing() {
    let src = r#"
Interface IOuter(Of T)
    Interface IInner(Of T)
        Function Combine(o As T) As String
    End Interface
End Interface

Class Impl
    Implements IOuter(Of String).IInner(Of Integer)
    Public Function Combine(o As Integer) As String Implements IOuter(Of String).IInner(Of Integer).Combine
        Return "IntegerVal_" & o
    End Function
End Class

Module Program
    Sub Main()
        Dim impl As IOuter(Of String).IInner(Of Integer) = New Impl()
        Console.WriteLine(impl.Combine(77))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntegerVal_77"]);
}
