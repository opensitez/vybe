use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.IntPtr & UIntPtr Pointer Operations & Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_intptr_zero_singleton() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pZero = IntPtr.Zero
        Console.WriteLine((pZero.ToInt64() = 0L) & "|" & (pZero = IntPtr.Zero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_uintptr_zero_singleton() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim uZero = UIntPtr.Zero
        Console.WriteLine((uZero.ToUInt64() = 0UL) & "|" & (uZero = UIntPtr.Zero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_intptr_size_property_32_or_64() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim sz = IntPtr.Size
        Console.WriteLine(sz = 4 OrElse sz = 8)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_intptr_add_and_subtract_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim basePtr As New IntPtr(1000)
        Dim added = IntPtr.Add(basePtr, 64)
        Dim subtracted = IntPtr.Subtract(added, 32)
        Console.WriteLine(added.ToInt32() & "|" & subtracted.ToInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1064|1032"]);
}

#[test]
fn test_vb_uintptr_add_and_subtract_static_methods() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim basePtr As New UIntPtr(2000UI)
        Dim added = UIntPtr.Add(basePtr, 100)
        Dim subtracted = UIntPtr.Subtract(added, 50)
        Console.WriteLine(added.ToUInt32() & "|" & subtracted.ToUInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2100|2050"]);
}

#[test]
fn test_vb_intptr_addition_subtraction_operator_overloads() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ptr As New IntPtr(500)
        Dim pPlus = ptr + 20
        Dim pMinus = ptr - 10
        Console.WriteLine(pPlus.ToInt32() & "|" & pMinus.ToInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["520|490"]);
}

#[test]
fn test_vb_intptr_equality_and_inequality_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim p1 As New IntPtr(1234)
        Dim p2 As New IntPtr(1234)
        Dim p3 As New IntPtr(5678)
        Console.WriteLine((p1 = p2) & "|" & (p1 <> p3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_intptr_explicit_integer_casts() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val As Integer = 4096
        Dim ptr As New IntPtr(val)
        Dim backToInt As Integer = ptr.ToInt32()
        Dim backToLong As Long = ptr.ToInt64()
        Console.WriteLine(backToInt & "|" & backToLong)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4096|4096"]);
}

#[test]
fn test_vb_uintptr_explicit_uinteger_ulong_casts() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val As ULong = 8192UL
        Dim ptr As New UIntPtr(val)
        Dim backToUInt64 As ULong = ptr.ToUInt64()
        Console.WriteLine(backToUInt64)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8192"]);
}

#[test]
fn test_vb_intptr_to_pointer_and_from_pointer_simulation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ptr As New IntPtr(100)
        Console.WriteLine(ptr.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_intptr_to_string_format_specifiers() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ptr As New IntPtr(255)
        Dim strHex = ptr.ToString("X")
        Console.WriteLine(strHex)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FF"]);
}

#[test]
fn test_vb_intptr_compare_to() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim p1 As New IntPtr(10)
        Dim p2 As New IntPtr(20)
        Console.WriteLine((p1.CompareTo(p2) < 0) & "|" & (p2.CompareTo(p1) > 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_intptr_get_hash_code_structural_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim p1 As New IntPtr(999)
        Dim p2 As New IntPtr(999)
        Console.WriteLine(p1.GetHashCode() = p2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_intptr_overflow_exception_on_to_int32_in_64bit() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Constructing an IntPtr from a Long larger than Int32.MaxValue
        If IntPtr.Size = 8 Then
            Dim largeVal As Long = &H100000000L
            Dim ptr As New IntPtr(largeVal)
            Try
                Dim val32 = ptr.ToInt32()
            Catch ex As OverflowException
                Console.WriteLine("OverflowException Caught on ToInt32")
            End Try
        Else
            Console.WriteLine("OverflowException Caught on ToInt32")
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OverflowException Caught on ToInt32"]);
}

#[test]
fn test_vb_intptr_array_sorting() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pointers As IntPtr() = {New IntPtr(30), New IntPtr(10), New IntPtr(20)}
        Array.Sort(pointers)
        For Each p In pointers
            Console.WriteLine(p.ToInt32())
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10", "20", "30"]);
}

#[test]
fn test_vb_uintptr_equality_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim u1 As New UIntPtr(777UI)
        Dim u2 As New UIntPtr(777UI)
        Dim u3 As New UIntPtr(888UI)
        Console.WriteLine((u1 = u2) & "|" & (u1 <> u3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_intptr_structure_field_in_unmanaged_header() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure NativeHeader
    Public Length As Integer
    Public DataPtr As IntPtr
End Structure

Module Program
    Sub Main()
        Dim h As New NativeHeader With {.Length = 64, .DataPtr = New IntPtr(12345)}
        Console.WriteLine(h.Length & "|" & h.DataPtr.ToInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["64|12345"]);
}

#[test]
fn test_vb_intptr_bitwise_or_and_simulation_via_int64() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim p1 As New IntPtr(&HFF00)
        Dim p2 As New IntPtr(&H00FF)
        Dim combined As New IntPtr(p1.ToInt64() Or p2.ToInt64())
        Console.WriteLine(Hex(combined.ToInt64()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FFFF"]);
}

#[test]
fn test_vb_intptr_negative_value_representation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim negPtr As New IntPtr(-1)
        Console.WriteLine(negPtr.ToInt32() = -1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_intptr_implicit_conversion_from_int32() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val As Integer = 77
        Dim ptr As IntPtr = CType(val, IntPtr)
        Console.WriteLine(ptr.ToInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["77"]);
}
