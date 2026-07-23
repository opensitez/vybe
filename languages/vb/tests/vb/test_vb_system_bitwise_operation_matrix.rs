use super::helpers::run_vb;

#[test]
fn bitwise_operations_integer_and_or_xor_identity() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Integer = &HF0
        Dim b As Integer = &H0F

        Console.WriteLine((a And b) = 0)
        Console.WriteLine((a Or b) = &HFF)
        Console.WriteLine((a Xor b) = &HFF)
        Console.WriteLine((a Xor (a Or b)) = b)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn bitwise_operations_shift_integers_are_masked() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim base As Integer = 1
        Console.WriteLine(base << 3)
        Console.WriteLine(base << 0)

        Dim shifted As Integer = 16 >> 1
        Console.WriteLine(shifted)
        Console.WriteLine((16 >> 4))
        Console.WriteLine((16 << 4))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8", "1", "8", "1", "256"]);
}

#[test]
fn bitwise_operations_negative_and_not_behavior() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = -1
        Dim y As Integer = x And 3
        Dim z As Integer = x Or 0
        Dim n As Integer = Not 0

        Console.WriteLine(y)
        Console.WriteLine(z)
        Console.WriteLine(n = -1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "-1", "True"]);
}

#[test]
fn bitwise_operations_byte_level_membership() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim mask As Byte = &HAA
        Dim hasA As Boolean = (mask And &H80) = &H80
        Dim has4 As Boolean = (mask And &H10) = &H10
        Dim onlyLow4 As Byte = mask And &HF0

        Console.WriteLine(mask)
        Console.WriteLine(hasA)
        Console.WriteLine(has4)
        Console.WriteLine(onlyLow4)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["170", "True", "False", "160"]);
}

#[test]
fn bitwise_operations_unsigned_compatibility() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim left As UInteger = 1UI
        Dim right As UInteger = &H80000000UI

        Dim combined As UInteger = left Or right
        Console.WriteLine(combined = CUInt(&H80000001))
        Console.WriteLine((combined And right) = right)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn bitwise_operations_bit_flags_from_boolean_set() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim read As Integer = 1
        Dim write As Integer = 2
        Dim execute As Integer = 4

        Dim flags As Integer = 0
        flags = flags Or read
        flags = flags Or execute

        Console.WriteLine(flags And read)
        Console.WriteLine(flags And write)
        Console.WriteLine(flags And (read Or execute))
        Console.WriteLine(flags = (read Or execute))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "0", "5", "True"]);
}

#[test]
fn bitwise_operations_rotate_like_pattern_with_shifts() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 0x12345678
        Dim highNibble As Integer = (x And &HF0000000) >> 28
        Dim lowNibble As Integer = x And &HF

        Console.WriteLine(highNibble)
        Console.WriteLine(lowNibble)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "8"]);
}
