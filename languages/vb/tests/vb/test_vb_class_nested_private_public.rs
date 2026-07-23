use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Nested Classes (Public, Private, Protected, Friend)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_nested_class_public_instantiation() {
    let src = r#"
Class Outer
    Public Class Inner
        Public Function GetValue() As String
            Return "Outer.Inner"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim obj As New Outer.Inner()
        Console.WriteLine(obj.GetValue())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Outer.Inner"]);
}

#[test]
fn test_vb_nested_class_private_accessed_by_outer() {
    let src = r#"
Class Outer
    Private Class PrivateInner
        Public Data As String = "SecretData"
    End Class

    Public Function GetInnerData() As String
        Dim inner As New PrivateInner()
        Return inner.Data
    End Function
End Class

Module Program
    Sub Main()
        Dim o As New Outer()
        Console.WriteLine(o.GetInnerData())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SecretData"]);
}

#[test]
fn test_vb_nested_class_accessing_outer_private_static_member() {
    let src = r#"
Class Outer
    Private Shared OuterSecret As String = "TopSecret"

    Public Class Inner
        Public Function ReadSecret() As String
            Return OuterSecret
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim inObj As New Outer.Inner()
        Console.WriteLine(inObj.ReadSecret())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TopSecret"]);
}

#[test]
fn test_vb_nested_class_protected_access_in_subclass() {
    let src = r#"
Class BaseOuter
    Protected Class ProtectedInner
        Public Message As String = "ProtectedInner"
    End Class
End Class

Class DerivedOuter
    Inherits BaseOuter
    Public Function GetProtectedMessage() As String
        Dim inner As New ProtectedInner()
        Return inner.Message
    End Function
End Class

Module Program
    Sub Main()
        Dim d As New DerivedOuter()
        Console.WriteLine(d.GetProtectedMessage())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ProtectedInner"]);
}

#[test]
fn test_vb_nested_class_deeply_nested_three_levels() {
    let src = r#"
Class Level1
    Public Class Level2
        Public Class Level3
            Public Shared Function Hello() As String
                Return "Level 3 Hello"
            End Function
        End Class
    End Class
End Class

Module Program
    Sub Main()
        Console.WriteLine(Level1.Level2.Level3.Hello())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Level 3 Hello"]);
}

#[test]
fn test_vb_nested_class_inheriting_from_outer_class() {
    let src = r#"
Class BaseClass
    Public Overridable Function Identify() As String
        Return "BaseClass"
    End Function

    Public Class NestedDerived
        Inherits BaseClass
        Public Overrides Function Identify() As String
            Return "NestedDerived"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim obj As BaseClass = New BaseClass.NestedDerived()
        Console.WriteLine(obj.Identify())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NestedDerived"]);
}

#[test]
fn test_vb_nested_class_outer_inheriting_from_nested_class() {
    let src = r#"
Class Outer
    Public Class InnerBase
        Public VirtualMsg As String = "InnerBaseMsg"
    End Class
End Class

Class SubOuter
    Inherits Outer.InnerBase
End Class

Module Program
    Sub Main()
        Dim s As New SubOuter()
        Console.WriteLine(s.VirtualMsg)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["InnerBaseMsg"]);
}

#[test]
fn test_vb_nested_enum_inside_class() {
    let src = r#"
Class NetworkConnection
    Public Enum State
        Disconnected
        Connecting
        Connected
    End Enum

    Public ConnectionState As State = State.Disconnected
End Class

Module Program
    Sub Main()
        Dim conn As New NetworkConnection()
        conn.ConnectionState = NetworkConnection.State.Connected
        Console.WriteLine(conn.ConnectionState.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Connected"]);
}

#[test]
fn test_vb_nested_interface_inside_class() {
    let src = r#"
Class Control
    Public Interface ICommandHandler
        Sub ExecuteCommand(cmd As String)
    End Interface

    Class ButtonHandler
        Implements ICommandHandler
        Public Sub ExecuteCommand(cmd As String) Implements ICommandHandler.ExecuteCommand
            Console.WriteLine("Button Command: " & cmd)
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Dim h As Control.ICommandHandler = New Control.ButtonHandler()
        h.ExecuteCommand("Click")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Button Command: Click"]);
}

#[test]
fn test_vb_nested_struct_inside_class() {
    let src = r#"
Class Graph
    Public Structure Node
        Public ID As Integer
        Public Label As String
        Public Sub New(id As Integer, label As String)
            Me.ID = id : Me.Label = label
        End Sub
    End Structure
End Class

Module Program
    Sub Main()
        Dim n As New Graph.Node(1, "Root")
        Console.WriteLine(n.ID & ":" & n.Label)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:Root"]);
}

#[test]
fn test_vb_nested_generic_class_inside_non_generic_class() {
    let src = r#"
Class Storage
    Public Class Cache(Of T)
        Private item As T
        Public Sub SetItem(val As T) : item = val : End Sub
        Public Function GetItem() As T : Return item : End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim c As New Storage.Cache(Of String)()
        c.SetItem("CachedData")
        Console.WriteLine(c.GetItem())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CachedData"]);
}

#[test]
fn test_vb_nested_non_generic_class_inside_generic_class() {
    let src = r#"
Class OuterList(Of T)
    Public Class Node
        Public Element As T
        Public NextNode As Node
        Public Sub New(e As T)
            Element = e
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Dim node1 As New OuterList(Of Integer).Node(10)
        Dim node2 As New OuterList(Of Integer).Node(20)
        node1.NextNode = node2
        Console.WriteLine(node1.Element & "->" & node1.NextNode.Element)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10->20"]);
}

#[test]
fn test_vb_nested_class_constructor_chaining_with_outer_instance() {
    let src = r#"
Class Document
    Public Title As String
    Public Sub New(t As String)
        Title = t
    End Sub

    Public Class Header
        Private parentDoc As Document
        Public Sub New(doc As Document)
            parentDoc = doc
        End Sub
        Public Function GetTitle() As String
            Return "Header of " & parentDoc.Title
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim doc As New Document("Report")
        Dim h As New Document.Header(doc)
        Console.WriteLine(h.GetTitle())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Header of Report"]);
}

#[test]
fn test_vb_nested_class_shadowing_outer_member_name() {
    let src = r#"
Class Outer
    Public Shared Value As Integer = 100

    Public Class Inner
        Public Shared Value As Integer = 200
        Public Shared Function PrintBoth() As String
            Return Value & "|" & Outer.Value
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Console.WriteLine(Outer.Inner.PrintBoth())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200|100"]);
}

#[test]
fn test_vb_nested_delegate_declaration() {
    let src = r#"
Class Processor
    Public Delegate Sub ProgressHandler(percent As Integer)

    Public Sub Run(handler As ProgressHandler)
        handler(50)
        handler(100)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Processor()
        p.Run(Sub(pct) Console.WriteLine("Progress: " & pct & "%"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Progress: 50%", "Progress: 100%"]);
}

#[test]
fn test_vb_nested_class_shared_constructor_execution() {
    let src = r#"
Class Outer
    Public Class Inner
        Public Shared Message As String
        Shared Sub New()
            Message = "Initialized Shared Inner"
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Console.WriteLine(Outer.Inner.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Initialized Shared Inner"]);
}

#[test]
fn test_vb_nested_class_friend_internal_accessibility() {
    let src = r#"
Class Parent
    Friend Class InternalHelper
        Public Shared Function Help() As String
            Return "Internal Help"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Console.WriteLine(Parent.InternalHelper.Help())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Internal Help"]);
}

#[test]
fn test_vb_nested_class_private_protected_mix() {
    let src = r#"
Class Base
    Protected Private Class Hidden
        Public ReadOnly Info As String = "HiddenInfo"
    End Class
    Public Function ReadHidden() As String
        Dim h As New Hidden()
        Return h.Info
    End Function
End Class

Module Program
    Sub Main()
        Dim b As New Base()
        Console.WriteLine(b.ReadHidden())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["HiddenInfo"]);
}

#[test]
fn test_vb_nested_class_reflection_type_name() {
    let src = r#"
Class ParentClass
    Public Class ChildClass
    End Class
End Class

Module Program
    Sub Main()
        Dim t = GetType(ParentClass.ChildClass)
        Console.WriteLine(t.Name & "|" & t.DeclaringType.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ChildClass|ParentClass"]);
}

#[test]
fn test_vb_nested_class_factory_pattern() {
    let src = r#"
Interface Product
    Function GetName() As String
End Interface

Class Factory
    Private Class ProductA
        Implements Product
        Public Function GetName() As String Implements Product.GetName : Return "ProdA" : End Function
    End Class

    Private Class ProductB
        Implements Product
        Public Function GetName() As String Implements Product.GetName : Return "ProdB" : End Function
    End Class

    Public Shared Function Create(type As String) As Product
        If type = "A" Then Return New ProductA()
        Return New ProductB()
    End Function
End Class

Module Program
    Sub Main()
        Dim p1 = Factory.Create("A")
        Dim p2 = Factory.Create("B")
        Console.WriteLine(p1.GetName() & "&" & p2.GetName())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ProdA&ProdB"]);
}
