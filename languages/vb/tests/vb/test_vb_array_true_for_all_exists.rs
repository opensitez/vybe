use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array.TrueForAll & Array.Exists Predicate Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_trueforall_all_match() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim evens As Integer() = {2, 4, 6, 8}
        Dim result As Boolean = Array.TrueForAll(evens, Function(n) n Mod 2 = 0)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_trueforall_one_fails() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {2, 4, 5, 8}
        Dim result As Boolean = Array.TrueForAll(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_array_trueforall_empty_array_is_vacuously_true() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Dim result As Boolean = Array.TrueForAll(empty, Function(n) n > 100)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_at_least_one_matches() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 3, 5, 8, 9}
        Dim hasEven As Boolean = Array.Exists(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(hasEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_none_matches() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 3, 5, 7, 9}
        Dim hasEven As Boolean = Array.Exists(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(hasEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_array_exists_empty_array_is_false() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Dim result As Boolean = Array.Exists(empty, Function(n) True)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_array_trueforall_string_length() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim words As String() = {"cat", "dog", "bat"}
        Dim allThreeChars As Boolean = Array.TrueForAll(words, Function(w) w.Length = 3)
        Console.WriteLine(allThreeChars)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_string_contains() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim files As String() = {"doc.txt", "data.csv", "image.png"}
        Dim hasCsv As Boolean = Array.Exists(files, Function(f) f.EndsWith(".csv"))
        Console.WriteLine(hasCsv)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_trueforall_complex_object_properties() {
    let src = r#"
Imports System

Class Account
    Public Balance As Decimal
    Public Sub New(b As Decimal)
        Balance = b
    End Sub
End Class

Module Program
    Sub Main()
        Dim accs As Account() = {New Account(100D), New Account(250D), New Account(50D)}
        Dim allPositive As Boolean = Array.TrueForAll(accs, Function(a) a.Balance > 0D)
        Dim allWealthy As Boolean = Array.TrueForAll(accs, Function(a) a.Balance >= 200D)
        Console.WriteLine(allPositive & "|" & allWealthy)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_array_exists_complex_object() {
    let src = r#"
Imports System

Class Account
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module Program
    Sub Main()
        Dim accs As Account() = {New Account("Alice"), New Account("Bob")}
        Dim hasBob As Boolean = Array.Exists(accs, Function(a) a.Name = "Bob")
        Console.WriteLine(hasBob)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_trueforall_short_circuits() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim calls As Integer = 0
        Dim numbers As Integer() = {10, -5, 20, 30}
        ' Should fail at second item (-5) and short-circuit
        Dim allPos As Boolean = Array.TrueForAll(numbers, Function(n)
            calls += 1
            Return n > 0
        End Function)
        Console.WriteLine(allPos & "|calls=" & calls)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|calls=2"]);
}

#[test]
fn test_vb_array_exists_short_circuits() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim calls As Integer = 0
        Dim numbers As Integer() = {10, -5, 20, 30}
        ' Should succeed at second item (-5) and short-circuit
        Dim hasNeg As Boolean = Array.Exists(numbers, Function(n)
            calls += 1
            Return n < 0
        End Function)
        Console.WriteLine(hasNeg & "|calls=" & calls)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|calls=2"]);
}

#[test]
fn test_vb_array_trueforall_struct_array() {
    let src = r#"
Imports System

Structure Item
    Public ID As Integer
    Public Sub New(id As Integer)
        Me.ID = id
    End Sub
End Structure

Module Program
    Sub Main()
        Dim items As Item() = {New Item(1), New Item(2), New Item(3)}
        Dim allValid As Boolean = Array.TrueForAll(items, Function(i) i.ID > 0)
        Console.WriteLine(allValid)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_struct_array() {
    let src = r#"
Imports System

Structure Item
    Public ID As Integer
    Public Sub New(id As Integer)
        Me.ID = id
    End Sub
End Structure

Module Program
    Sub Main()
        Dim items As Item() = {New Item(1), New Item(2), New Item(3)}
        Dim hasTwo As Boolean = Array.Exists(items, Function(i) i.ID = 2)
        Console.WriteLine(hasTwo)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_trueforall_nullable_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim items As Nullable(Of Integer)() = {10, 20, 30}
        Dim allHasValue As Boolean = Array.TrueForAll(items, Function(i) i.HasValue)
        Console.WriteLine(allHasValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_nullable_array_with_null() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim items As Nullable(Of Integer)() = {10, Nothing, 30}
        Dim hasNull As Boolean = Array.Exists(items, Function(i) Not i.HasValue)
        Console.WriteLine(hasNull)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_trueforall_enum_array() {
    let src = r#"
Imports System

Enum Status
    Active
    Pending
    Inactive
End Enum

Module Program
    Sub Main()
        Dim states As Status() = {Status.Active, Status.Active}
        Dim allActive As Boolean = Array.TrueForAll(states, Function(s) s = Status.Active)
        Console.WriteLine(allActive)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_enum_array() {
    let src = r#"
Imports System

Enum Status
    Active
    Pending
    Inactive
End Enum

Module Program
    Sub Main()
        Dim states As Status() = {Status.Active, Status.Pending}
        Dim hasInactive As Boolean = Array.Exists(states, Function(s) s = Status.Inactive)
        Console.WriteLine(hasInactive)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_array_trueforall_method_address_of() {
    let src = r#"
Imports System

Module Validator
    Public Function IsPositive(n As Integer) As Boolean
        Return n > 0
    End Function
End Module

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4}
        Dim result As Boolean = Array.TrueForAll(numbers, AddressOf Validator.IsPositive)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_array_exists_method_address_of() {
    let src = r#"
Imports System

Module Validator
    Public Function IsZero(n As Integer) As Boolean
        Return n = 0
    End Function
End Module

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 0, 4}
        Dim result As Boolean = Array.Exists(numbers, AddressOf Validator.IsZero)
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
