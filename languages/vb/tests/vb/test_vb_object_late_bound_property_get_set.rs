use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Object Late-Bound Property Access & Indexers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_late_bound_property_get_set() {
    let src = r#"
Module Program
    Class Account
        Public Property Balance As Decimal
    End Class

    Sub Main()
        Dim obj As Object = New Account()
        obj.Balance = 500.50D
        Dim b As Decimal = CDec(obj.Balance)
        Console.WriteLine(b)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["500.50"]);
}

#[test]
fn test_vb_late_bound_field_get_set() {
    let src = r#"
Module Program
    Class Settings
        Public AppName As String
    End Class

    Sub Main()
        Dim obj As Object = New Settings()
        obj.AppName = "VybeEngine"
        Dim name As String = CStr(obj.AppName)
        Console.WriteLine(name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VybeEngine"]);
}

#[test]
fn test_vb_late_bound_read_only_property_get() {
    let src = r#"
Module Program
    Class VersionInfo
        Public ReadOnly Property SystemVersion As String
            Get
                Return "1.0.0"
            End Get
        End Property
    End Class

    Sub Main()
        Dim obj As Object = New VersionInfo()
        Dim ver As String = CStr(obj.SystemVersion)
        Console.WriteLine(ver)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.0.0"]);
}

#[test]
fn test_vb_late_bound_write_only_property_set() {
    let src = r#"
Module Program
    Class Vault
        Private secretKey As String
        Public WriteOnly Property Password As String
            Set(value As String)
                secretKey = value
            End Set
        End Property
        Public Function GetKeyLength() As Integer
            Return If(secretKey IsNot Nothing, secretKey.Length, 0)
        End Function
    End Class

    Sub Main()
        Dim v As New Vault()
        Dim obj As Object = v
        obj.Password = "SuperSecret"
        Console.WriteLine(v.GetKeyLength())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11"]);
}

#[test]
fn test_vb_late_bound_default_indexer_get_set() {
    let src = r#"
Module Program
    Class CustomDictionary
        Private storage As New System.Collections.Generic.Dictionary(Of String, String)()
        Default Public Property Item(key As String) As String
            Get
                Return storage(key)
            End Get
            Set(value As String)
                storage(key) = value
            End Set
        End Property
    End Class

    Sub Main()
        Dim obj As Object = New CustomDictionary()
        obj("host") = "localhost"
        Dim host As String = CStr(obj("host"))
        Console.WriteLine(host)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["localhost"]);
}

#[test]
fn test_vb_late_bound_array_element_get_set() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30}
        Dim obj As Object = numbers
        obj(1) = 99
        Dim v As Integer = CInt(obj(1))
        Console.WriteLine(v & "|" & String.Join(",", numbers))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99|10,99,30"]);
}

#[test]
fn test_vb_late_bound_2d_array_element_get_set() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim obj As Object = grid
        obj(1, 0) = 300
        Dim val As Integer = CInt(obj(1, 0))
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["300"]);
}

#[test]
fn test_vb_late_bound_property_compound_assignment() {
    let src = r#"
Module Program
    Class Counter
        Public Property Value As Integer = 5
    End Class

    Sub Main()
        Dim obj As Object = New Counter()
        obj.Value += 10
        Dim v As Integer = CInt(obj.Value)
        Console.WriteLine(v)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_late_bound_property_string_concatenation_assignment() {
    let src = r#"
Module Program
    Class StringBuilderHolder
        Public Property Content As String = "Hello"
    End Class

    Sub Main()
        Dim obj As Object = New StringBuilderHolder()
        obj.Content &= " World"
        Dim res As String = CStr(obj.Content)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_late_bound_property_struct_type() {
    let src = r#"
Module Program
    Structure Size2D
        Public Width As Integer
        Public Height As Integer
    End Structure

    Class Frame
        Public Property Bounds As Size2D
    End Class

    Sub Main()
        Dim obj As Object = New Frame()
        obj.Bounds = New Size2D With {.Width = 1920, .Height = 1080}
        Dim sz As Size2D = CType(obj.Bounds, Size2D)
        Console.WriteLine(sz.Width & "x" & sz.Height)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1920x1080"]);
}

#[test]
fn test_vb_late_bound_property_enum_type() {
    let src = r#"
Enum PriorityLevel
    Low = 1
    Critical = 10
End Enum

Module Program
    Class TaskItem
        Public Property Priority As PriorityLevel
    End Class

    Sub Main()
        Dim obj As Object = New TaskItem()
        obj.Priority = PriorityLevel.Critical
        Dim p As PriorityLevel = CType(obj.Priority, PriorityLevel)
        Console.WriteLine(p.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Critical"]);
}

#[test]
fn test_vb_late_bound_property_set_read_only_throws() {
    let src = r#"
Imports System

Module Program
    Class ConstantValue
        Public ReadOnly Property ID As Integer = 100
    End Class

    Sub Main()
        Dim obj As Object = New ConstantValue()
        Try
            obj.ID = 200
        Catch ex As Exception
            Console.WriteLine("Property Set on ReadOnly Property Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Property Set on ReadOnly Property Caught"]
    );
}

#[test]
fn test_vb_late_bound_property_get_write_only_throws() {
    let src = r#"
Imports System

Module Program
    Class Sink
        Public WriteOnly Property Data As String
            Set(value As String)
            End Set
        End Property
    End Class

    Sub Main()
        Dim obj As Object = New Sink()
        Try
            Dim val = obj.Data
        Catch ex As Exception
            Console.WriteLine("Property Get on WriteOnly Property Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Property Get on WriteOnly Property Caught"]
    );
}

#[test]
fn test_vb_late_bound_property_type_mismatch_assignment_throws() {
    let src = r#"
Imports System

Module Program
    Class StrictHolder
        Public Property Count As Integer
    End Class

    Sub Main()
        Dim obj As Object = New StrictHolder()
        Try
            obj.Count = "NotAnIntegerString"
        Catch ex As Exception
            Console.WriteLine("Invalid Cast in Late Bound Property Set Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Invalid Cast in Late Bound Property Set Caught"]
    );
}

#[test]
fn test_vb_late_bound_nested_property_get() {
    let src = r#"
Module Program
    Class Company
        Public Property Address As AddressInfo
    End Class

    Class AddressInfo
        Public Property City As String
    End Class

    Sub Main()
        Dim c As New Company With {.Address = New AddressInfo With {.City = "Tokyo"}}
        Dim obj As Object = c
        Console.WriteLine(CStr(obj.Address.City))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Tokyo"]);
}

#[test]
fn test_vb_late_bound_multi_parameter_indexed_property() {
    let src = r#"
Module Program
    Class Matrix2D
        Private data(1, 1) As Double
        Default Public Property Item(row As Integer, col As Integer) As Double
            Get
                Return data(row, col)
            End Get
            Set(value As Double)
                data(row, col) = value
            End Set
        End Property
    End Class

    Sub Main()
        Dim obj As Object = New Matrix2D()
        obj(0, 1) = 7.7
        Dim val As Double = CDbl(obj(0, 1))
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["7.7"]);
}

#[test]
fn test_vb_late_bound_static_property_via_instance() {
    let src = r#"
Module Program
    Class AppGlobal
        Public Shared Property Counter As Integer = 42
    End Class

    Sub Main()
        Dim obj As Object = New AppGlobal()
        ' Late bound access via instance dispatches to shared property!
        Console.WriteLine(CInt(obj.Counter))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_late_bound_field_increment() {
    let src = r#"
Module Program
    Class Stats
        Public Hits As Long = 1000
    End Class

    Sub Main()
        Dim obj As Object = New Stats()
        obj.Hits += 1L
        Console.WriteLine(CLng(obj.Hits))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1001"]);
}

#[test]
fn test_vb_late_bound_interface_property_access() {
    let src = r#"
Module Program
    Interface INamed
        Property Name As String
    End Interface

    Class NamedItem
        Implements INamed
        Public Property Name As String Implements INamed.Name
    End Class

    Sub Main()
        Dim obj As Object = New NamedItem()
        obj.Name = "InterfaceProperty"
        Console.WriteLine(CStr(obj.Name))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["InterfaceProperty"]);
}

#[test]
fn test_vb_late_bound_property_null_assignment() {
    let src = r#"
Module Program
    Class NullableHolder
        Public Property Text As String = "Initial"
    End Class

    Sub Main()
        Dim obj As Object = New NullableHolder()
        obj.Text = Nothing
        Console.WriteLine(obj.Text Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
