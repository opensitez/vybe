use super::helpers::run_vb;

#[test]
fn bitwise_and_masks_bits() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(12 And 10)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8"]);
}

#[test]
fn bitwise_or_sets_bits() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(12 Or 3)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["15"]);
}

#[test]
fn bitwise_xor_toggles_bits() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(10 Xor 12)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["6"]);
}

#[test]
fn bitwise_not_byte_is_inverse_within_width() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(CByte(Not CByte(&HF)) )
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["240"]);
}

#[test]
fn bitwise_left_shift_multiplies_by_power_of_two() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(1 << 4)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["16"]);
}

#[test]
fn bitwise_right_shift_divides_by_power_of_two() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(64 >> 3)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8"]);
}

#[test]
fn bitwise_right_shift_preserves_sign_for_negative_values() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine((-8) >> 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["-4"]);
}

#[test]
fn bitwise_compound_and_updates_in_place() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim flags As Integer = &H0F
        flags = flags And &H05
        Console.WriteLine(flags)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn bitwise_compound_or_updates_in_place() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim flags As Integer = &H08
        flags = flags Or &H03
        Console.WriteLine(flags)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["11"]);
}

#[test]
fn bitwise_compound_xor_updates_in_place() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim flags As Integer = &H0F
        flags = flags Xor &H05
        Console.WriteLine(flags)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn bitwise_bitmask_test_with_flag_set() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim flags As Integer = &H0A
        Console.WriteLine((flags And &H02) <> 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn bitwise_bitmask_test_without_flag_clear() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim flags As Integer = &H0A
        Console.WriteLine((flags And &H01) <> 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn bitwise_combines_negation_and_mask() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Integer = 1
        Console.WriteLine((Not value) And 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn bitwise_roundtrip_right_then_left_shift_keeps_parity_even() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Integer = 9
        value = value << 2
        value = value >> 2
        Console.WriteLine(value = 9)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
