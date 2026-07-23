use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Threading.Interlocked Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_interlocked_increment_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim counter As Integer = 10
        Dim newVal = Interlocked.Increment(counter)
        Console.WriteLine(newVal & "|" & counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11|11"]);
}

#[test]
fn test_vb_interlocked_decrement_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim counter As Integer = 10
        Dim newVal = Interlocked.Decrement(counter)
        Console.WriteLine(newVal & "|" & counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9|9"]);
}

#[test]
fn test_vb_interlocked_increment_long() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim counter As Long = 100L
        Dim newVal = Interlocked.Increment(counter)
        Console.WriteLine(newVal & "|" & counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101|101"]);
}

#[test]
fn test_vb_interlocked_add_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim total As Integer = 50
        Dim newVal = Interlocked.Add(total, 25)
        Console.WriteLine(newVal & "|" & total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["75|75"]);
}

#[test]
fn test_vb_interlocked_add_negative_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim total As Integer = 50
        Dim newVal = Interlocked.Add(total, -10)
        Console.WriteLine(newVal & "|" & total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["40|40"]);
}

#[test]
fn test_vb_interlocked_exchange_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim target As Integer = 100
        Dim oldVal = Interlocked.Exchange(target, 200)
        Console.WriteLine("Old: " & oldVal & " | New: " & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: 100 | New: 200"]);
}

#[test]
fn test_vb_interlocked_exchange_reference_type() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim ref1 As String = "Original"
        Dim oldVal = Interlocked.Exchange(ref1, "Replaced")
        Console.WriteLine("Old: " & oldVal & " | New: " & ref1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: Original | New: Replaced"]);
}

#[test]
fn test_vb_interlocked_compare_exchange_match() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim target As Integer = 50
        ' Replace with 99 IF current value equals 50
        Dim oldVal = Interlocked.CompareExchange(target, 99, 50)
        Console.WriteLine("Old: " & oldVal & " | Current: " & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: 50 | Current: 99"]);
}

#[test]
fn test_vb_interlocked_compare_exchange_no_match() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim target As Integer = 50
        ' Attempt to replace with 99 IF current value equals 100 (which it doesn't!)
        Dim oldVal = Interlocked.CompareExchange(target, 99, 100)
        Console.WriteLine("Old: " & oldVal & " | Current: " & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: 50 | Current: 50"]);
}

#[test]
fn test_vb_interlocked_compare_exchange_reference_type_match() {
    let src = r#"
Imports System.Threading

Class Document
    Public Title As String
    Public Sub New(t As String) : Title = t : End Sub
End Class

Module Program
    Sub Main()
        Dim doc1 As New Document("D1")
        Dim doc2 As New Document("D2")
        Dim target As Document = doc1

        Dim oldVal = Interlocked.CompareExchange(target, doc2, doc1)
        Console.WriteLine(Object.ReferenceEquals(oldVal, doc1) & "|" & target.Title)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|D2"]);
}

#[test]
fn test_vb_interlocked_read_64bit_long() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim val As Long = 9876543210L
        Dim current = Interlocked.Read(val)
        Console.WriteLine(current)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9876543210"]);
}

#[test]
fn test_vb_interlocked_exchange_single_float() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim target As Single = 1.5F
        Dim oldVal = Interlocked.Exchange(target, 3.5F)
        Console.WriteLine(oldVal & "|" & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.5|3.5"]);
}

#[test]
fn test_vb_interlocked_exchange_double_float() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim target As Double = 10.5
        Dim oldVal = Interlocked.Exchange(target, 20.5)
        Console.WriteLine(oldVal & "|" & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10.5|20.5"]);
}

#[test]
fn test_vb_interlocked_compare_exchange_generic_reference() {
    let src = r#"
Imports System.Threading

Module Program
    Private Function CompareExchangeGeneric(Of T As Class)(ByRef location As T, value As T, comparand As T) As T
        Return Interlocked.CompareExchange(location, value, comparand)
    End Function

    Sub Main()
        Dim s1 As String = "Alpha"
        Dim s2 As String = "Beta"
        Dim target As String = s1
        Dim oldVal = CompareExchangeGeneric(target, s2, s1)
        Console.WriteLine(oldVal & "|" & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alpha|Beta"]);
}

#[test]
fn test_vb_interlocked_spinlock_lock_free_counter() {
    let src = r#"
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim counter As Integer = 0
        Dim tasks(9) As Task
        For i As Integer = 0 To 9
            tasks(i) = Task.Run(Sub()
                For j As Integer = 1 To 100
                    Interlocked.Increment(counter)
                Next
            End Sub)
        Next
        Task.WaitAll(tasks)
        Console.WriteLine(counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1000"]);
}

#[test]
fn test_vb_interlocked_or_bitwise_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim flags As Integer = 1 ' 0001
        Dim oldVal = Interlocked.Or(flags, 4) ' 0100 -> 0101 (5)
        Console.WriteLine("Old: " & oldVal & " | New: " & flags)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: 1 | New: 5"]);
}

#[test]
fn test_vb_interlocked_and_bitwise_integer() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim flags As Integer = 7 ' 0111
        Dim oldVal = Interlocked.And(flags, 3) ' 0011 -> 0011 (3)
        Console.WriteLine("Old: " & oldVal & " | New: " & flags)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: 7 | New: 3"]);
}

#[test]
fn test_vb_interlocked_memory_barrier() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim data As Integer = 42
        Interlocked.MemoryBarrier()
        Console.WriteLine(data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_interlocked_compare_exchange_null_check() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim target As String = Nothing
        Dim oldVal = Interlocked.CompareExchange(target, "Initialized", Nothing)
        Console.WriteLine((oldVal Is Nothing) & "|" & target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Initialized"]);
}

#[test]
fn test_vb_interlocked_exchange_intptr() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim ptr1 As New IntPtr(1000)
        Dim ptr2 As New IntPtr(2000)
        Dim oldPtr = Interlocked.Exchange(ptr1, ptr2)
        Console.WriteLine(oldPtr.ToInt32() & "|" & ptr1.ToInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1000|2000"]);
}
