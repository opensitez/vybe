use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interface Implements & Inheritance Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_interface_implements_single_method() {
    let src = r#"
Interface IPrintable
    Sub Print()
End Interface

Class Document
    Implements IPrintable
    Public Sub Print() Implements IPrintable.Print
        Console.WriteLine("Printing Document")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IPrintable = New Document()
        p.Print()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Printing Document"]);
}

#[test]
fn test_vb_interface_implements_multiple_interfaces() {
    let src = r#"
Interface IReader
    Function ReadData() As String
End Interface

Interface IWriter
    Sub WriteData(data As String)
End Interface

Class StorageStream
    Implements IReader, IWriter
    Private buffer As String = ""
    Public Function ReadData() As String Implements IReader.ReadData
        Return buffer
    End Function
    Public Sub WriteData(data As String) Implements IWriter.WriteData
        buffer = data
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New StorageStream()
        Dim w As IWriter = s
        w.WriteData("Hello World")
        Dim r As IReader = s
        Console.WriteLine(r.ReadData())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_interface_explicit_name_aliasing() {
    let src = r#"
Interface ILogger
    Sub Log(msg As String)
End Interface

Class FileLogger
    Implements ILogger
    Public Sub RecordMessage(msg As String) Implements ILogger.Log
        Console.WriteLine("LOG: " & msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As ILogger = New FileLogger()
        l.Log("System Started")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LOG: System Started"]);
}

#[test]
fn test_vb_interface_property_implementation() {
    let src = r#"
Interface INamed
    Property Name As String
End Interface

Class User
    Implements INamed
    Public Property Name As String Implements INamed.Name
End Class

Module Program
    Sub Main()
        Dim n As INamed = New User() With {.Name = "Alice"}
        Console.WriteLine(n.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_interface_read_only_property() {
    let src = r#"
Interface IIdentifiable
    ReadOnly Property Id As Integer
End Interface

Class Item
    Implements IIdentifiable
    Private _id As Integer
    Public Sub New(id As Integer)
        _id = id
    End Sub
    Public ReadOnly Property Id As Integer Implements IIdentifiable.Id
        Get
            Return _id
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim item As IIdentifiable = New Item(42)
        Console.WriteLine(item.Id)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_interface_write_only_property() {
    let src = r#"
Interface IPasswordReceiver
    WriteOnly Property Password As String
End Interface

Class Service
    Implements IPasswordReceiver
    Private _pwd As String
    Public WriteOnly Property Password As String Implements IPasswordReceiver.Password
        Set(value As String)
            _pwd = value
        End Set
    End Property
    Public Function Verify(p As String) As Boolean
        Return _pwd = p
    End Function
End Class

Module Program
    Sub Main()
        Dim s As New Service()
        Dim r As IPasswordReceiver = s
        r.Password = "Secret123"
        Console.WriteLine(s.Verify("Secret123"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_interface_event_implementation() {
    let src = r#"
Imports System

Interface INotifier
    Event Notified(msg As String)
    Sub Trigger(msg As String)
End Interface

Class Notifier
    Implements INotifier
    Public Event Notified(msg As String) Implements INotifier.Notified
    Public Sub Trigger(msg As String) Implements INotifier.Trigger
        RaiseEvent Notified(msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As INotifier = New Notifier()
        AddHandler n.Notified, Sub(m) Console.WriteLine("RECEIVED: " & m)
        n.Trigger("Alert")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["RECEIVED: Alert"]);
}

#[test]
fn test_vb_interface_generic_method() {
    let src = r#"
Interface IConverter
    Function Convert(Of TInput, TOutput)(input As TInput) As TOutput
End Interface

Class StringConverter
    Implements IConverter
    Public Function Convert(Of TInput, TOutput)(input As TInput) As TOutput Implements IConverter.Convert
        Return CType(CObj(input.ToString()), TOutput)
    End Function
End Class

Module Program
    Sub Main()
        Dim c As IConverter = New StringConverter()
        Dim res As String = c.Convert(Of Integer, String)(100)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_interface_generic_interface() {
    let src = r#"
Interface IRepository(Of T)
    Sub Add(entity As T)
    Function GetById(id As Integer) As T
End Interface

Class ProductRepository
    Implements IRepository(Of String)
    Private item As String = ""
    Public Sub Add(entity As String) Implements IRepository(Of String).Add
        item = entity
    End Sub
    Public Function GetById(id As Integer) As String Implements IRepository(Of String).GetById
        Return item
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As IRepository(Of String) = New ProductRepository()
        repo.Add("Laptop")
        Console.WriteLine(repo.GetById(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Laptop"]);
}

#[test]
fn test_vb_interface_inheritance_chain() {
    let src = r#"
Interface IBase
    Sub MethodA()
End Interface

Interface IDerived
    Inherits IBase
    Sub MethodB()
End Interface

Class Implementation
    Implements IDerived
    Public Sub MethodA() Implements IBase.MethodA
        Console.WriteLine("MethodA")
    End Sub
    Public Sub MethodB() Implements IDerived.MethodB
        Console.WriteLine("MethodB")
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As IDerived = New Implementation()
        d.MethodA()
        d.MethodB()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["MethodA", "MethodB"]);
}

#[test]
fn test_vb_interface_multiple_interface_inheritance() {
    let src = r#"
Interface IReadable
    Sub Read()
End Interface

Interface IWritable
    Sub Write()
End Interface

Interface IStreamable
    Inherits IReadable, IWritable
    Sub Flush()
End Interface

Class MemoryStreamHandler
    Implements IStreamable
    Public Sub Read() Implements IReadable.Read : Console.WriteLine("Read") : End Sub
    Public Sub Write() Implements IWritable.Write : Console.WriteLine("Write") : End Sub
    Public Sub Flush() Implements IStreamable.Flush : Console.WriteLine("Flush") : End Sub
End Class

Module Program
    Sub Main()
        Dim s As IStreamable = New MemoryStreamHandler()
        s.Read()
        s.Write()
        s.Flush()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Read", "Write", "Flush"]);
}

#[test]
fn test_vb_interface_implementation_in_abstract_class() {
    let src = r#"
Interface IService
    Sub Execute()
End Interface

MustInherit Class BaseService
    Implements IService
    Public MustOverride Sub Execute() Implements IService.Execute
End Class

Class CustomService
    Inherits BaseService
    Public Overrides Sub Execute()
        Console.WriteLine("Custom Execution")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IService = New CustomService()
        s.Execute()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Custom Execution"]);
}

#[test]
fn test_vb_interface_single_class_multiple_methods_same_interface() {
    let src = r#"
Interface IMathOps
    Function Add(a As Integer, b As Integer) As Integer
    Function Add(a As Double, b As Double) As Double
End Interface

Class Calculator
    Implements IMathOps
    Public Function Add(a As Integer, b As Integer) As Integer Implements IMathOps.Add
        Return a + b
    End Function
    Public Function Add(a As Double, b As Double) As Double Implements IMathOps.Add
        Return a + b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As IMathOps = New Calculator()
        Console.WriteLine(calc.Add(5, 10) & "|" & calc.Add(2.5, 3.5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15|6"]);
}

#[test]
fn test_vb_interface_disposable_pattern() {
    let src = r#"
Imports System

Class Resource
    Implements IDisposable
    Public Disposed As Boolean = False
    Public Sub Dispose() Implements IDisposable.Dispose
        Disposed = True
        Console.WriteLine("Disposed")
    End Sub
End Class

Module Program
    Sub Main()
        Using r As New Resource()
            Console.WriteLine("Using Resource")
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Using Resource", "Disposed"]);
}

#[test]
fn test_vb_interface_struct_implements_interface() {
    let src = r#"
Interface IFormattableValue
    Function Format() As String
End Interface

Structure Currency
    Implements IFormattableValue
    Public Amount As Decimal
    Public Sub New(amt As Decimal)
        Amount = amt
    End Sub
    Public Function Format() As String Implements IFormattableValue.Format
        Return "$" & Amount
    End Function
End Structure

Module Program
    Sub Main()
        Dim c As IFormattableValue = New Currency(49.99D)
        Console.WriteLine(c.Format())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["$49.99"]);
}

#[test]
fn test_vb_interface_co_variance_out_keyword() {
    let src = r#"
Interface IProducer(Of Out T)
    Function Produce() As T
End Interface

Class StringProducer
    Implements IProducer(Of String)
    Public Function Produce() As String Implements IProducer(Of String).Produce
        Return "Produced String"
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IProducer(Of Object) = New StringProducer()
        Console.WriteLine(p.Produce().ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Produced String"]);
}

#[test]
fn test_vb_interface_contra_variance_in_keyword() {
    let src = r#"
Interface IConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class ObjectConsumer
    Implements IConsumer(Of Object)
    Public Sub Consume(item As Object) Implements IConsumer(Of Object).Consume
        Console.WriteLine("Consumed: " & item.ToString())
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As IConsumer(Of String) = New ObjectConsumer()
        c.Consume("Data String")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Consumed: Data String"]);
}

#[test]
fn test_vb_interface_combining_interfaces_with_generic_constraints() {
    let src = r#"
Interface IValidatable
    Function IsValid() As Boolean
End Interface

Class Processor(Of T As IValidatable)
    Public Function Process(item As T) As String
        If item.IsValid() Then
            Return "Valid"
        Else
            Return "Invalid"
        End If
    End Function
End Class

Class FormInput
    Implements IValidatable
    Public Input As String
    Public Sub New(i As String)
        Input = i
    End Sub
    Public Function IsValid() As Boolean Implements IValidatable.IsValid
        Return Not String.IsNullOrEmpty(Input)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Processor(Of FormInput)()
        Console.WriteLine(p.Process(New FormInput("OK")))
        Console.WriteLine(p.Process(New FormInput("")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Valid", "Invalid"]);
}

#[test]
fn test_vb_interface_private_interface_nested_in_class() {
    let src = r#"
Class Outer
    Private Interface IInternal
        Sub InternalWork()
    End Interface
    Private Class Inner
        Implements IInternal
        Public Sub InternalWork() Implements IInternal.InternalWork
            Console.WriteLine("Internal Work Done")
        End Sub
    End Class
    Public Sub Run()
        Dim i As IInternal = New Inner()
        i.InternalWork()
    End Sub
End Class

Module Program
    Sub Main()
        Dim o As New Outer()
        o.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Internal Work Done"]);
}

#[test]
fn test_vb_interface_array_of_interfaces() {
    let src = r#"
Interface IShape
    Function Area() As Double
End Interface

Class Circle
    Implements IShape
    Public Radius As Double
    Public Sub New(r As Double) : Radius = r : End Sub
    Public Function Area() As Double Implements IShape.Area
        Return Math.PI * Radius * Radius
    End Function
End Class

Class Square
    Implements IShape
    Public Side As Double
    Public Sub New(s As Double) : Side = s : End Sub
    Public Function Area() As Double Implements IShape.Area
        Return Side * Side
    End Function
End Class

Module Program
    Sub Main()
        Dim shapes As IShape() = {New Circle(10), New Square(10)}
        Console.WriteLine(Math.Round(shapes(0).Area(), 2) & "|" & shapes(1).Area())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["314.16|100"]);
}
