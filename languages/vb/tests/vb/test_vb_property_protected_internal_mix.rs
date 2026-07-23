use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Property Access Modifiers (Protected, Friend, Private Set)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_property_public_get_private_set() {
    let src = r#"
Class Entity
    Public Property Id As Integer { Get; Private Set; }
    Public Sub New(id As Integer)
        Me.Id = id
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Entity(101)
        Console.WriteLine(e.Id)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101"]);
}

#[test]
fn test_vb_property_public_get_protected_set() {
    let src = r#"
Class BaseAccount
    Public Property Balance As Decimal { Get; Protected Set; }
    Public Sub New(bal As Decimal)
        Balance = bal
    End Sub
End Class

Class CheckingAccount
    Inherits BaseAccount
    Public Sub New(bal As Decimal)
        MyBase.New(bal)
    End Sub
    Public Sub Deposit(amount As Decimal)
        Balance += amount
    End Sub
End Class

Module Program
    Sub Main()
        Dim acc As New CheckingAccount(100D)
        acc.Deposit(50D)
        Console.WriteLine(acc.Balance)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150"]);
}

#[test]
fn test_vb_property_protected_getter_and_setter() {
    let src = r#"
Class BaseService
    Protected Property SecretKey As String = "BaseKey"
End Class

Class CustomService
    Inherits BaseService
    Public Function GetKey() As String
        Return SecretKey
    End Function
End Class

Module Program
    Sub Main()
        Dim s As New CustomService()
        Console.WriteLine(s.GetKey())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BaseKey"]);
}

#[test]
fn test_vb_property_friend_internal_access() {
    let src = r#"
Class ModuleConfig
    Friend Property Mode As String = "Debug"
End Class

Module Program
    Sub Main()
        Dim cfg As New ModuleConfig()
        Console.WriteLine(cfg.Mode)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Debug"]);
}

#[test]
fn test_vb_property_protected_friend_access() {
    let src = r#"
Class BaseNode
    Protected Friend Property ConnectionString As String = "Server=localhost"
End Class

Module Program
    Sub Main()
        Dim node As New BaseNode()
        Console.WriteLine(node.ConnectionString)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Server=localhost"]);
}

#[test]
fn test_vb_property_private_protected_access_in_derived() {
    let src = r#"
Class BaseFramework
    Private Protected Property InternalTag As String = "TagV1"
    Public Function GetTag() As String
        Return InternalTag
    End Function
End Class

Module Program
    Sub Main()
        Dim f As New BaseFramework()
        Console.WriteLine(f.GetTag())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TagV1"]);
}

#[test]
fn test_vb_property_public_get_friend_set() {
    let src = r#"
Class Package
    Public Property Status As String { Get; Friend Set; } = "Created"
End Class

Module Program
    Sub Main()
        Dim p As New Package()
        p.Status = "Dispatched"
        Console.WriteLine(p.Status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Dispatched"]);
}

#[test]
fn test_vb_property_override_access_modifier_compatibility() {
    let src = r#"
Class BaseClass
    Public Overridable Property Message As String
        Get
            Return "BaseMsg"
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class DerivedClass
    Inherits BaseClass
    Public Overrides Property Message As String
        Get
            Return "DerivedMsg"
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim b As BaseClass = New DerivedClass()
        Console.WriteLine(b.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DerivedMsg"]);
}

#[test]
fn test_vb_property_abstract_mustoverride_protected() {
    let src = r#"
MustInherit Class AbstractModel
    Protected MustOverride Property Code As Integer
    Public Function ReadCode() As Integer
        Return Code
    End Function
End Class

Class ConcreteModel
    Inherits AbstractModel
    Protected Overrides Property Code As Integer = 999
End Class

Module Program
    Sub Main()
        Dim m As AbstractModel = New ConcreteModel()
        Console.WriteLine(m.ReadCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999"]);
}

#[test]
fn test_vb_property_read_only_with_backing_field_private() {
    let src = r#"
Class Product
    Private _price As Decimal
    Public ReadOnly Property Price As Decimal
        Get
            Return _price
        End Get
    End Property
    Public Sub New(p As Decimal)
        _price = p
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Product(29.99D)
        Console.WriteLine(p.Price)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["29.99"]);
}

#[test]
fn test_vb_property_write_only_private_backing_field() {
    let src = r#"
Class SystemToken
    Private _token As String
    Public WriteOnly Property Token As String
        Set(value As String)
            _token = "ENC_" & value
        End Set
    End Property
    Public Function ValidateToken(input As String) As Boolean
        Return _token = "ENC_" & input
    End Function
End Class

Module Program
    Sub Main()
        Dim st As New SystemToken()
        st.Token = "Pass123"
        Console.WriteLine(st.ValidateToken("Pass123"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_property_expression_bodied_public_get() {
    let src = r#"
Class Circle
    Public Property Radius As Double
    Public ReadOnly Property Area As Double => Math.PI * Radius * Radius
    Public Sub New(r As Double)
        Radius = r
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Circle(10)
        Console.WriteLine(Math.Round(c.Area, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["314.16"]);
}

#[test]
fn test_vb_property_shared_public_get_private_set() {
    let src = r#"
Class GlobalCounter
    Public Shared Property TotalCount As Integer { Get; Private Set; } = 0
    Public Shared Sub Increment()
        TotalCount += 1
    End Sub
End Class

Module Program
    Sub Main()
        GlobalCounter.Increment()
        GlobalCounter.Increment()
        Console.WriteLine(GlobalCounter.TotalCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_property_interface_explicit_implementation_private_in_class() {
    let src = r#"
Interface IInternalData
    ReadOnly Property Data As String
End Interface

Class SecretProvider
    Implements IInternalData
    Private ReadOnly Property Data As String Implements IInternalData.Data
        Get
            Return "SecretDataValue"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim provider As IInternalData = New SecretProvider()
        Console.WriteLine(provider.Data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SecretDataValue"]);
}

#[test]
fn test_vb_property_struct_public_get_private_set() {
    let src = r#"
Structure StructPoint
    Public Property X As Integer { Get; Private Set; }
    Public Property Y As Integer { Get; Private Set; }
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim p As New StructPoint(10, 20)
        Console.WriteLine(p.X & "," & p.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_property_generic_class_private_setter() {
    let src = r#"
Class Box(Of T)
    Public Property Item As T { Get; Private Set; }
    Public Sub New(val As T)
        Item = val
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New Box(Of String)("Contents")
        Console.WriteLine(b.Item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Contents"]);
}

#[test]
fn test_vb_property_protected_indexer_access() {
    let src = r#"
Class BaseArray
    Private arr(2) As Integer
    Protected Default Property Item(idx As Integer) As Integer
        Get
            Return arr(idx)
        End Get
        Set(value As Integer)
            arr(idx) = value
        End Set
    End Property
End Class

Class CustomArray
    Inherits BaseArray
    Public Sub SetValue(idx As Integer, val As Integer)
        MyBase.Item(idx) = val
    End Sub
    Public Function GetValue(idx As Integer) As Integer
        Return MyBase.Item(idx)
    End Function
End Class

Module Program
    Sub Main()
        Dim ca As New CustomArray()
        ca.SetValue(0, 99)
        Console.WriteLine(ca.GetValue(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_property_reflection_get_set_access_rights() {
    let src = r#"
Class Sample
    Public Property Text As String { Get; Private Set; } = "Init"
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Sample).GetProperty("Text")
        Console.WriteLine((prop.GetMethod IsNot Nothing) & "|" & (prop.SetMethod IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_property_custom_get_set_validation() {
    let src = r#"
Class AgeValidator
    Private _age As Integer
    Public Property Age As Integer
        Get
            Return _age
        End Get
        Set(value As Integer)
            If value >= 0 Then _age = value Else _age = 0
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim a As New AgeValidator()
        a.Age = 25
        Console.WriteLine(a.Age)
        a.Age = -5
        Console.WriteLine(a.Age)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["25", "0"]);
}

#[test]
fn test_vb_property_auto_property_initializer_with_private_set() {
    let src = r#"
Class SystemDefaults
    Public Property MaxRetries As Integer { Get; Private Set; } = 3
End Class

Module Program
    Sub Main()
        Dim sd As New SystemDefaults()
        Console.WriteLine(sd.MaxRetries)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}
