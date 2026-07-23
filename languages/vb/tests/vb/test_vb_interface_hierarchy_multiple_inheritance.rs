use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Interface Hierarchy, Multi-Inheritance & Disambiguation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_interface_multiple_inheritance_method_ambiguity_resolved() {
    let src = r#"
Interface IX
    Sub Print()
End Interface

Interface IY
    Sub Print()
End Interface

Class Implementation
    Implements IX, IY
    Private Sub PrintX() Implements IX.Print
        Console.WriteLine("IX Print")
    End Sub
    Private Sub PrintY() Implements IY.Print
        Console.WriteLine("IY Print")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New Implementation()
        Dim x As IX = obj
        Dim y As IY = obj
        x.Print()
        y.Print()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IX Print", "IY Print"]);
}

#[test]
fn test_vb_interface_single_method_implements_multiple_interface_members() {
    let src = r#"
Interface IX
    Sub Common()
End Interface

Interface IY
    Sub Common()
End Interface

Class SharedImpl
    Implements IX, IY
    Public Sub Common() Implements IX.Common, IY.Common
        Console.WriteLine("Shared Common Execution")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New SharedImpl()
        Dim x As IX = obj
        Dim y As IY = obj
        x.Common()
        y.Common()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Shared Common Execution", "Shared Common Execution"]
    );
}

#[test]
fn test_vb_interface_diamond_inheritance_hierarchy() {
    let src = r#"
Interface IBase
    Sub BaseMethod()
End Interface

Interface ILeft
    Inherits IBase
    Sub LeftMethod()
End Interface

Interface IRight
    Inherits IBase
    Sub RightMethod()
End Interface

Interface IDiamond
    Inherits ILeft, IRight
    Sub DiamondMethod()
End Interface

Class DiamondImpl
    Implements IDiamond
    Public Sub BaseMethod() Implements IBase.BaseMethod
        Console.WriteLine("BaseMethod")
    End Sub
    Public Sub LeftMethod() Implements ILeft.LeftMethod
        Console.WriteLine("LeftMethod")
    End Sub
    Public Sub RightMethod() Implements IRight.RightMethod
        Console.WriteLine("RightMethod")
    End Sub
    Public Sub DiamondMethod() Implements IDiamond.DiamondMethod
        Console.WriteLine("DiamondMethod")
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As IDiamond = New DiamondImpl()
        d.BaseMethod()
        d.LeftMethod()
        d.RightMethod()
        d.DiamondMethod()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["BaseMethod", "LeftMethod", "RightMethod", "DiamondMethod"]
    );
}

#[test]
fn test_vb_interface_property_conflict_disambiguation() {
    let src = r#"
Interface IAlpha
    ReadOnly Property Value As Integer
End Interface

Interface IBeta
    ReadOnly Property Value As String
End Interface

Class Component
    Implements IAlpha, IBeta
    Public ReadOnly Property AlphaValue As Integer Implements IAlpha.Value
        Get
            Return 42
        End Get
    End Property
    Public ReadOnly Property BetaValue As String Implements IBeta.Value
        Get
            Return "FortyTwo"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim c As New Component()
        Dim a As IAlpha = c
        Dim b As IBeta = c
        Console.WriteLine(a.Value & "|" & b.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|FortyTwo"]);
}

#[test]
fn test_vb_interface_event_conflict_disambiguation() {
    let src = r#"
Imports System

Interface IEventA
    Event OnEvent As EventHandler
End Interface

Interface IEventB
    Event OnEvent As EventHandler
End Interface

Class Dispatcher
    Implements IEventA, IEventB
    Public Event EventA As EventHandler Implements IEventA.OnEvent
    Public Event EventB As EventHandler Implements IEventB.OnEvent

    Public Sub RaiseA()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
    Public Sub RaiseB()
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New Dispatcher()
        Dim ea As IEventA = d
        Dim eb As IEventB = d

        AddHandler ea.OnEvent, Sub(s, e) Console.WriteLine("A Raised")
        AddHandler eb.OnEvent, Sub(s, e) Console.WriteLine("B Raised")

        d.RaiseA()
        d.RaiseB()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A Raised", "B Raised"]);
}

#[test]
fn test_vb_interface_reimplementation_in_derived_class() {
    let src = r#"
Interface IPrintable
    Sub Print()
End Interface

Class Parent
    Implements IPrintable
    Public Overridable Sub Print() Implements IPrintable.Print
        Console.WriteLine("Parent Print")
    End Sub
End Class

Class Child
    Inherits Parent
    Implements IPrintable
    Public Overrides Sub Print() Implements IPrintable.Print
        Console.WriteLine("Child Reimplemented Print")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IPrintable = New Child()
        p.Print()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Child Reimplemented Print"]);
}

#[test]
fn test_vb_interface_shadowing_interface_member() {
    let src = r#"
Interface IBase
    Sub Render()
End Interface

Interface IDerived
    Inherits IBase
    Shadows Sub Render()
End Interface

Class Window
    Implements IDerived
    Public Sub BaseRender() Implements IBase.Render
        Console.WriteLine("IBase Render")
    End Sub
    Public Sub DerivedRender() Implements IDerived.Render
        Console.WriteLine("IDerived Render")
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As New Window()
        Dim b As IBase = w
        Dim d As IDerived = w
        b.Render()
        d.Render()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IBase Render", "IDerived Render"]);
}

#[test]
fn test_vb_interface_generic_interface_inheritance_specialization() {
    let src = r#"
Interface IRepository(Of T)
    Sub Add(item As T)
End Interface

Interface ICustomerRepository
    Inherits IRepository(Of String)
    Function GetCustomerName(id As Integer) As String
End Interface

Class CustomerService
    Implements ICustomerRepository
    Private customer As String
    Public Sub Add(item As String) Implements IRepository(Of String).Add
        customer = item
    End Sub
    Public Function GetCustomerName(id As Integer) As String Implements ICustomerRepository.GetCustomerName
        Return customer
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As ICustomerRepository = New CustomerService()
        repo.Add("Alice")
        Console.WriteLine(repo.GetCustomerName(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_interface_casting_between_sibling_interfaces() {
    let src = r#"
Interface ISource
    Sub SourceAction()
End Interface

Interface ITarget
    Sub TargetAction()
End Interface

Class DualAction
    Implements ISource, ITarget
    Public Sub SourceAction() Implements ISource.SourceAction
        Console.WriteLine("Source")
    End Sub
    Public Sub TargetAction() Implements ITarget.TargetAction
        Console.WriteLine("Target")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As ISource = New DualAction()
        Dim t As ITarget = CType(s, ITarget)
        t.TargetAction()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Target"]);
}

#[test]
fn test_vb_interface_is_operator_runtime_type_check() {
    let src = r#"
Interface ISupportFastSearch
End Interface

Class FastDatabase
    Implements ISupportFastSearch
End Class

Class SlowDatabase
End Class

Module Program
    Sub Main()
        Dim db1 As Object = New FastDatabase()
        Dim db2 As Object = New SlowDatabase()
        Console.WriteLine((TypeOf db1 Is ISupportFastSearch) & "|" & (TypeOf db2 Is ISupportFastSearch))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_interface_trycast_safe_conversion() {
    let src = r#"
Interface IDisposableResource
    Sub CleanUp()
End Interface

Class SafeResource
    Implements IDisposableResource
    Public Sub CleanUp() Implements IDisposableResource.CleanUp
        Console.WriteLine("Cleaned Up")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New SafeResource()
        Dim res = TryCast(obj, IDisposableResource)
        If res IsNot Nothing Then
            res.CleanUp()
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cleaned Up"]);
}

#[test]
fn test_vb_interface_trycast_returns_nothing_for_non_implementer() {
    let src = r#"
Interface IRunnable
    Sub Run()
End Interface

Class NonRunnable
End Class

Module Program
    Sub Main()
        Dim obj As Object = New NonRunnable()
        Dim r = TryCast(obj, IRunnable)
        Console.WriteLine(r Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_interface_default_interface_method_override_in_subinterface() {
    let src = r#"
Interface ILogger
    Sub Log(msg As String)
End Interface

Interface IAdvancedLogger
    Inherits ILogger
    Sub Log(msg As String, severity As Integer)
End Interface

Class CustomLogger
    Implements IAdvancedLogger
    Public Sub Log(msg As String) Implements ILogger.Log
        Console.WriteLine("Basic: " & msg)
    End Sub
    Public Sub Log(msg As String, severity As Integer) Implements IAdvancedLogger.Log
        Console.WriteLine("Advanced [" & severity & "]: " & msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As IAdvancedLogger = New CustomLogger()
        l.Log("System Start")
        l.Log("Critical Failure", 5)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Basic: System Start", "Advanced [5]: Critical Failure"]
    );
}

#[test]
fn test_vb_interface_struct_implementing_multiple_interfaces() {
    let src = r#"
Interface IX : Function GetX() As Integer : End Interface
Interface IY : Function GetY() As Integer : End Interface

Structure Point2D
    Implements IX, IY
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
    Public Function GetX() As Integer Implements IX.GetX : Return X : End Function
    Public Function GetY() As Integer Implements IY.GetY : Return Y : End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Point2D(10, 20)
        Dim ix As IX = p
        Dim iy As IY = p
        Console.WriteLine(ix.GetX() & "," & iy.GetY())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_interface_indexer_multiple_overloads() {
    let src = r#"
Interface IIndexable
    Default Property Item(key As String) As String
    Default Property Item(index As Integer) As String
End Interface

Class DictionaryAdapter
    Implements IIndexable
    Public Property Item(key As String) As String Implements IIndexable.Item
        Get
            Return "Key_" & key
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property Item(index As Integer) As String Implements IIndexable.Item
        Get
            Return "Idx_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim idx As IIndexable = New DictionaryAdapter()
        Console.WriteLine(idx("name") & "|" & idx(42))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Key_name|Idx_42"]);
}

#[test]
fn test_vb_interface_hierarchy_deep_inheritance_resolution() {
    let src = r#"
Interface IA : Sub ActA() : End Interface
Interface IB : Inherits IA : Sub ActB() : End Interface
Interface IC : Inherits IB : Sub ActC() : End Interface

Class DeepImpl
    Implements IC
    Public Sub ActA() Implements IA.ActA : Console.WriteLine("A") : End Sub
    Public Sub ActB() Implements IB.ActB : Console.WriteLine("B") : End Sub
    Public Sub ActC() Implements IC.ActC : Console.WriteLine("C") : End Sub
End Class

Module Program
    Sub Main()
        Dim c As IC = New DeepImpl()
        c.ActA() : c.ActB() : c.ActC()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A", "B", "C"]);
}

#[test]
fn test_vb_interface_abstract_class_partial_implementation() {
    let src = r#"
Interface IFullService
    Sub ActionA()
    Sub ActionB()
End Interface

MustInherit Class PartialService
    Implements IFullService
    Public Sub ActionA() Implements IFullService.ActionA
        Console.WriteLine("ActionA Completed")
    End Sub
    Public MustOverride Sub ActionB() Implements IFullService.ActionB
End Class

Class CompleteService
    Inherits PartialService
    Public Overrides Sub ActionB()
        Console.WriteLine("ActionB Completed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IFullService = New CompleteService()
        s.ActionA()
        s.ActionB()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ActionA Completed", "ActionB Completed"]);
}

#[test]
fn test_vb_interface_multiple_interface_return_types() {
    let src = r#"
Interface IEntity
    ReadOnly Property ID As Integer
End Interface

Interface IAudit
    ReadOnly Property CreatedAt As String
End Interface

Class AuditedEntity
    Implements IEntity, IAudit
    Public ReadOnly Property ID As Integer Implements IEntity.ID
        Get
            Return 101
        End Get
    End Property
    Public ReadOnly Property CreatedAt As String Implements IAudit.CreatedAt
        Get
            Return "2025-01-01"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim ae As New AuditedEntity()
        Console.WriteLine(ae.ID & "|" & ae.CreatedAt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101|2025-01-01"]);
}

#[test]
fn test_vb_interface_delegates_as_interface_properties() {
    let src = r#"
Imports System

Interface ICallbackContainer
    Property Handler As Action(Of String)
End Interface

Class Worker
    Implements ICallbackContainer
    Public Property Handler As Action(Of String) Implements ICallbackContainer.Handler
    Public Sub Run()
        If Handler IsNot Nothing Then
            Handler("Finished Work")
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As New Worker()
        Dim c As ICallbackContainer = w
        c.Handler = Sub(msg) Console.WriteLine("Callback: " & msg)
        w.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Callback: Finished Work"]);
}

#[test]
fn test_vb_interface_empty_marker_interface() {
    let src = r#"
Interface ISerializableMarker
End Interface

Class DataPacket
    Implements ISerializableMarker
End Class

Module Program
    Sub Main()
        Dim p As Object = New DataPacket()
        Console.WriteLine(TypeOf p Is ISerializableMarker)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
