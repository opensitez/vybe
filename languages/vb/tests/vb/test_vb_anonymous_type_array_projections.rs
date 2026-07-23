use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Anonymous Types Array Projections & Key Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_anonymous_type_basic_properties() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Name = "Alice", .Age = 25}
        Console.WriteLine(obj.Name & " is " & obj.Age)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice is 25"]);
}

#[test]
fn test_vb_anonymous_type_key_property_immutable() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {Key .ID = 101, .Status = "Active"}
        ' Key properties participate in Equals & GetHashCode
        Console.WriteLine(obj.ID & "|" & obj.Status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101|Active"]);
}

#[test]
fn test_vb_anonymous_type_key_properties_equals_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim o1 = New With {Key .ID = 1, .Name = "A"}
        Dim o2 = New With {Key .ID = 1, .Name = "B"} ' Non-key Name ignored in Equals
        Dim o3 = New With {Key .ID = 2, .Name = "A"}
        Console.WriteLine(o1.Equals(o2) & "|" & o1.Equals(o3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_anonymous_type_array_implicit_type_inference() {
    let src = r#"
Module Program
    Sub Main()
        Dim people = {
            New With {.Name = "Alice", .Score = 90},
            New With {.Name = "Bob", .Score = 85}
        }
        For Each p In people
            Console.WriteLine(p.Name & ":" & p.Score)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice:90", "Bob:85"]);
}

#[test]
fn test_vb_anonymous_type_linq_select_projection() {
    let src = r#"
Imports System.Linq

Class Employee
    Public Property Name As String
    Public Property Salary As Double
    Public Sub New(n As String, s As Double) : Name = n : Salary = s : End Sub
End Class

Module Program
    Sub Main()
        Dim emps = {New Employee("Alice", 50000), New Employee("Bob", 60000)}
        Dim projected = From e In emps Select New With {.EmpName = e.Name, .AnnualSalary = e.Salary}
        For Each item In projected
            Console.WriteLine(item.EmpName & "=" & item.AnnualSalary)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice=50000", "Bob=60000"]);
}

#[test]
fn test_vb_anonymous_type_linq_group_by_projection() {
    let src = r#"
Imports System.Linq

Class Student
    Public Property Grade As Integer
    Public Property Name As String
    Public Sub New(g As Integer, n As String) : Grade = g : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim students = {New Student(10, "Alice"), New Student(10, "Bob"), New Student(11, "Charlie")}
        Dim groups = From s In students
                     Group s By s.Grade Into Group
                     Select New With {.Grade = Grade, .Count = Group.Count()}

        For Each g In groups
            Console.WriteLine("Grade " & g.Grade & " Count: " & g.Count)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Grade 10 Count: 2", "Grade 11 Count: 1"]);
}

#[test]
fn test_vb_anonymous_type_property_inferred_name_from_variable() {
    let src = r#"
Module Program
    Sub Main()
        Dim title As String = "Manager"
        Dim level As Integer = 3
        ' Inferred property names .title and .level
        Dim obj = New With {title, level}
        Console.WriteLine(obj.title & ":" & obj.level)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Manager:3"]);
}

#[test]
fn test_vb_anonymous_type_property_inferred_name_from_member_access() {
    let src = r#"
Class User
    Public Property Username As String = "AdminUser"
End Class

Module Program
    Sub Main()
        Dim u As New User()
        ' Inferred property name .Username
        Dim obj = New With {u.Username}
        Console.WriteLine(obj.Username)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AdminUser"]);
}

#[test]
fn test_vb_anonymous_type_to_string_representation() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.X = 10, .Y = 20}
        Dim str = obj.ToString()
        Console.WriteLine(str.Contains("X = 10") AndAlso str.Contains("Y = 20"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_nested_anonymous_objects() {
    let src = r#"
Module Program
    Sub Main()
        Dim person = New With {
            .Name = "Alice",
            .Address = New With {.City = "Seattle", .Zip = "98101"}
        }
        Console.WriteLine(person.Name & " in " & person.Address.City)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice in Seattle"]);
}

#[test]
fn test_vb_anonymous_type_mutable_property() {
    let src = r#"
Module Program
    Sub Main()
        ' In VB, non-Key properties of anonymous types are mutable!
        Dim item = New With {.Price = 10.0}
        item.Price = 15.5
        Console.WriteLine(item.Price)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15.5"]);
}

#[test]
fn test_vb_anonymous_type_hash_code_consistency() {
    let src = r#"
Module Program
    Sub Main()
        Dim o1 = New With {Key .Code = "A1"}
        Dim o2 = New With {Key .Code = "A1"}
        Console.WriteLine(o1.GetHashCode() = o2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_in_dictionary_lookup() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Object, String)()
        Dim keyObj = New With {Key .ID = 42}
        dict(keyObj) = "FoundData"
        Dim lookupObj = New With {Key .ID = 42}
        Console.WriteLine(dict(lookupObj))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FoundData"]);
}

#[test]
fn test_vb_anonymous_type_linq_join_projection() {
    let src = r#"
Imports System.Linq

Class Order
    Public Property OrderID As Integer
    Public Property CustomerID As Integer
    Public Sub New(o As Integer, c As Integer) : OrderID = o : CustomerID = c : End Sub
End Class

Class Customer
    Public Property CustomerID As Integer
    Public Property Name As String
    Public Sub New(c As Integer, n As String) : CustomerID = c : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim orders = {New Order(1, 101), New Order(2, 102)}
        Dim customers = {New Customer(101, "Alice"), New Customer(102, "Bob")}

        Dim joined = From o In orders
                     Join c In customers On o.CustomerID Equals c.CustomerID
                     Select New With {.OrderID = o.OrderID, .CustomerName = c.Name}

        For Each item In joined
            Console.WriteLine("No." & item.OrderID & " " & item.CustomerName)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["No.1 Alice", "No.2 Bob"]);
}

#[test]
fn test_vb_anonymous_type_with_lambda_expression_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj = New With {
            .Multiplier = 2,
            .Calc = Function(n As Integer) n * 2
        }
        Console.WriteLine(obj.Calc(10))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_anonymous_type_generic_type_args_discovery() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Str = "Text", .Num = 123}
        Dim props = obj.GetType().GetProperties()
        Console.WriteLine(props.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_anonymous_type_array_of_nested_anonymous_types() {
    let src = r#"
Module Program
    Sub Main()
        Dim items = {
            New With {.Meta = New With {.Tag = "T1"}},
            New With {.Meta = New With {.Tag = "T2"}}
        }
        Console.WriteLine(items(0).Meta.Tag & "|" & items(1).Meta.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["T1|T2"]);
}

#[test]
fn test_vb_anonymous_type_null_property_values() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Text = CType(Nothing, String), .Val = CType(Nothing, Nullable(Of Integer))}
        Console.WriteLine((obj.Text Is Nothing) & "|" & (Not obj.Val.HasValue))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_anonymous_type_enum_property() {
    let src = r#"
Enum Priority
    Low
    High
End Enum

Module Program
    Sub Main()
        Dim obj = New With {.Level = Priority.High}
        Console.WriteLine(obj.Level.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["High"]);
}

#[test]
fn test_vb_anonymous_type_tuple_property() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Coords = (10, 20)}
        Console.WriteLine(obj.Coords.Item1 & "," & obj.Coords.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}
