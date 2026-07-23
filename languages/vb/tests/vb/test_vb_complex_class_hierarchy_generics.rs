use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Deep Class Hierarchies, Generic Constraints & Interfaces
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_repository_pattern_crud() {
    let src = r#"
Imports System.Collections.Generic

Interface IEntity
    Property Id As Integer
End Interface

Class User
    Implements IEntity
    Public Property Id As Integer Implements IEntity.Id
    Public Property Name As String
End Class

Class Repository(Of T As {Class, IEntity, New})
    Private items As New List(Of T)()

    Public Sub Add(item As T)
        items.Add(item)
    End Sub

    Public Function GetById(id As Integer) As T
        For Each item In items
            If item.Id = id Then Return item
        Next
        Return Nothing
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As New Repository(Of User)()
        repo.Add(New User With {.Id = 1, .Name = "Alice"})
        repo.Add(New User With {.Id = 2, .Name = "Bob"})

        Dim u = repo.GetById(2)
        Console.WriteLine(u.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bob"]);
}

#[test]
fn test_vb_abstract_template_method_pattern() {
    let src = r#"
MustInherit Class DataProcessor
    Public Sub Process()
        Step1()
        Step2()
    End Sub

    Protected MustOverride Sub Step1()
        Protected MustOverride Sub Step2()
End Class

Class XmlProcessor
    Inherits DataProcessor
    Protected Overrides Sub Step1()
        Console.WriteLine("Xml Step 1")
    End Sub
    Protected Overrides Sub Step2()
        Console.WriteLine("Xml Step 2")
    End Sub
End Class

Module Program
    Sub Main()
        Dim proc As DataProcessor = New XmlProcessor()
        proc.Process()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Xml Step 1", "Xml Step 2"]);
}

#[test]
fn test_vb_generic_covariance_out_parameter() {
    let src = r#"
Interface IReadOnlyRepository(Of Out T)
    Function GetFirst() As T
End Interface

Class StringRepository
    Implements IReadOnlyRepository(Of String)
    Public Function GetFirst() As String Implements IReadOnlyRepository(Of String).GetFirst
        Return "CovariantResult"
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As IReadOnlyRepository(Of String) = New StringRepository()
        Dim objRepo As IReadOnlyRepository(Of Object) = repo ' Covariant assignment!
        Console.WriteLine(objRepo.GetFirst().ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CovariantResult"]);
}

#[test]
fn test_vb_generic_contravariance_in_parameter() {
    let src = r#"
Interface IReceiver(Of In T)
    Sub Receive(data As T)
End Interface

Class ObjectReceiver
    Implements IReceiver(Of Object)
    Public Sub Receive(data As Object) Implements IReceiver(Of Object).Receive
        Console.WriteLine("Received: " & data.ToString())
    End Sub
End Class

Module Program
    Sub Main()
        Dim objRec As IReceiver(Of Object) = New ObjectReceiver()
        Dim strRec As IReceiver(Of String) = objRec ' Contravariant assignment!
        strRec.Receive("ContravariantPayload")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Received: ContravariantPayload"]);
}

#[test]
fn test_vb_multiple_interface_inheritance_conflict_resolution() {
    let src = r#"
Interface ILoggerA
    Sub Log(msg As String)
End Interface

Interface ILoggerB
    Sub Log(msg As String)
End Interface

Class DualLogger
    Implements ILoggerA, ILoggerB

    Private Sub LogA(msg As String) Implements ILoggerA.Log
        Console.WriteLine("LoggerA: " & msg)
    End Sub

    Private Sub LogB(msg As String) Implements ILoggerB.Log
        Console.WriteLine("LoggerB: " & msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim dl As New DualLogger()
        Dim a As ILoggerA = dl
        Dim b As ILoggerB = dl
        a.Log("Message")
        b.Log("Message")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LoggerA: Message", "LoggerB: Message"]);
}

#[test]
fn test_vb_sealed_notinheritable_class_behavior() {
    let src = r#"
NotInheritable Class FinalConfig
    Public Property Version As Integer = 1
End Class

Module Program
    Sub Main()
        Dim fc As New FinalConfig()
        Console.WriteLine(fc.Version)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_nested_generic_class_instantiation() {
    let src = r#"
Class OuterTree(Of TKey)
    Class Node(Of TValue)
        Public Key As TKey
        Public Value As TValue
        Public Sub New(k As TKey, v As TValue)
            Key = k
            Value = v
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Dim node As New OuterTree(Of String).Node(Of Integer)("Key1", 100)
        Console.WriteLine(node.Key & "=" & node.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Key1=100"]);
}

#[test]
fn test_vb_generic_self_referencing_constraint() {
    let src = r#"
Imports System

Class ComparableBase(Of T As ComparableBase(Of T))
    Implements IComparable(Of T)
    Public Property Priority As Integer

    Public Function CompareTo(other As T) As Integer Implements IComparable(Of T).CompareTo
        Return Priority.CompareTo(other.Priority)
    End Function
End Class

Class DerivedItem
    Inherits ComparableBase(Of DerivedItem)
End Class

Module Program
    Sub Main()
        Dim item1 As New DerivedItem With {.Priority = 10}
        Dim item2 As New DerivedItem With {.Priority = 20}
        Console.WriteLine(item1.CompareTo(item2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_override_property_getter_setter_expansion() {
    let src = r#"
Class BaseProperty
    Public Overridable Property Value As Integer = 10
End Class

Class LoggedProperty
    Inherits BaseProperty

    Public Overrides Property Value As Integer
        Get
            Return MyBase.Value
        End Get
        Set(v As Integer)
            Console.WriteLine("Setting Value to " & v)
            MyBase.Value = v
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim p As BaseProperty = New LoggedProperty()
        p.Value = 42
        Console.WriteLine(p.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Setting Value to 42", "42"]);
}

#[test]
fn test_vb_generic_factory_new_constraint() {
    let src = r#"
Module Program
    Private Function CreateInstance(Of T As {Class, New})() As T
        Return New T()
    End Function

    Class Component
        Public ReadOnly Created As Boolean = True
    End Class

    Sub Main()
        Dim c = CreateInstance(Of Component)()
        Console.WriteLine(c.Created)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_class_hierarchy_shadows_keyword() {
    let src = r#"
Class Parent
    Public Sub Display()
        Console.WriteLine("Parent Display")
    End Sub
End Class

Class Child
    Inherits Parent
    Public Shadows Sub Display()
        Console.WriteLine("Child Display")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Child()
        Dim p As Parent = c
        c.Display()
        p.Display()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Child Display", "Parent Display"]);
}

#[test]
fn test_vb_interface_default_methods_implementation() {
    let src = r#"
Interface IService
    Sub Execute()
End Interface

Class DefaultService
    Implements IService
    Public Sub Execute() Implements IService.Execute
        Console.WriteLine("Executed Default Service")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IService = New DefaultService()
        s.Execute()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Executed Default Service"]);
}

#[test]
fn test_vb_generic_struct_method_invocations() {
    let src = r#"
Structure Pair(Of T1, T2)
    Public First As T1
    Public Second As T2
    Public Sub New(f As T1, s As T2)
        First = f
        Second = s
    End Sub
    Public Function Swap() As Pair(Of T2, T1)
        Return New Pair(Of T2, T1)(Second, First)
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of String, Integer)("Age", 30)
        Dim swapped = p.Swap()
        Console.WriteLine(swapped.First & "=" & swapped.Second)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30=Age"]);
}

#[test]
fn test_vb_class_hierarchy_virtual_event_raising() {
    let src = r#"
Imports System

Class BaseEmitter
    Public Event Notice As EventHandler
    Protected Overridable Sub OnNotice()
        RaiseEvent Notice(Me, EventArgs.Empty)
    End Sub
    Public Sub Fire()
        OnNotice()
    End Sub
End Class

Class InterceptEmitter
    Inherits BaseEmitter
    Protected Overrides Sub OnNotice()
        Console.WriteLine("Intercepted Before Fire")
        MyBase.OnNotice()
    End Sub
End Class

Module Program
    Sub Main()
        Dim ie As New InterceptEmitter()
        AddHandler ie.Notice, Sub(s, e) Console.WriteLine("Base Notice Fired")
        ie.Fire()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Intercepted Before Fire", "Base Notice Fired"]
    );
}

#[test]
fn test_vb_generic_enum_constraint_simulation() {
    let src = r#"
Imports System

Module Program
    Private Function GetEnumName(Of T As {Structure, System.IConvertible})(val As T) As String
        Return [Enum].GetName(GetType(T), val)
    End Function

    Enum Status
        Active = 1
    End Enum

    Sub Main()
        Dim name = GetEnumName(Status.Active)
        Console.WriteLine(name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Active"]);
}

#[test]
fn test_vb_class_hierarchy_multiple_constructors_mybase_new() {
    let src = r#"
Class BasePerson
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Class Employee
    Inherits BasePerson

    Public Salary As Decimal
    Public Sub New(n As String, s As Decimal)
        MyBase.New(n)
        Salary = s
    End Sub
End Class

Module Program
    Sub Main()
        Dim emp As New Employee("Alice", 75000D)
        Console.WriteLine(emp.Name & ":" & emp.Salary)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice:75000"]);
}

#[test]
fn test_vb_generic_delegate_invocation() {
    let src = r#"
Delegate Function Transform(Of T, R)(item As T) As R

Module Program
    Sub Main()
        Dim stringLen As Transform(Of String, Integer) = Function(s) s.Length
        Console.WriteLine(stringLen("VisualBasic"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11"]);
}

#[test]
fn test_vb_class_hierarchy_protected_friend_access() {
    let src = r#"
Class Parent
    Protected Friend Shared Sub InternalSharedLog()
        Console.WriteLine("Protected Friend Shared Log")
    End Sub
End Class

Module Program
    Sub Main()
        Parent.InternalSharedLog()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Protected Friend Shared Log"]);
}

#[test]
fn test_vb_interface_explicit_implementation_hiding() {
    let src = r#"
Interface ISecret
    Sub Reveal()
End Interface

Class Vault
    Implements ISecret
    Private Sub RevealSecret() Implements ISecret.Reveal
        Console.WriteLine("Secret Revealed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim v As New Vault()
        Dim s As ISecret = v
        s.Reveal()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Secret Revealed"]);
}

#[test]
fn test_vb_generic_class_static_field_per_type_instantiation() {
    let src = r#"
Class Counter(Of T)
    Public Shared Count As Integer = 0
End Class

Module Program
    Sub Main()
        Counter(Of String).Count = 100
        Counter(Of Integer).Count = 200
        Console.WriteLine(Counter(Of String).Count & "|" & Counter(Of Integer).Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|200"]);
}
