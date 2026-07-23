use super::helpers::run_vb;

#[test]
fn integer_arithmetic_matrix_add_sub_mul_identity() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values() As Integer = {-12, -3, 0, 1, 2, 5, 10, 17}

        Dim allGood As Boolean = True

        For Each a In values
            For Each b In values
                If (a + b - b <> a) Then
                    allGood = False
                End If
                If (a - b + b <> a) Then
                    allGood = False
                End If
                If (a * 1 <> a) Then
                    allGood = False
                End If
            Next
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn integer_arithmetic_matrix_division_mod_contracts() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim numerators() As Integer = {-8, -3, 2, 5, 17}
        Dim denominators() As Integer = {-3, 1, 2, 4}

        Dim allGood As Boolean = True

        For Each n In numerators
            For Each d In denominators
                Dim q As Integer = n \ d
                Dim r As Integer = n Mod d
                If (q * d + r <> n) Then
                    allGood = False
                End If
            Next
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn integer_arithmetic_matrix_bitwise_compositions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values() As Integer = {0, 1, 2, 3, 5, 8, 13}
        Dim allGood As Boolean = True

        For Each a In values
            If ((a << 1) >> 1 <> a) Then allGood = False
            If ((a >> 1) <= a) = False Then allGood = False
            If ((a Xor a) <> 0) Then allGood = False
            If ((a Or 0) <> a) Then allGood = False
            If ((a And 0) <> 0) Then allGood = False
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn integer_arithmetic_matrix_unary_and_sign() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values() As Integer = {0, 1, -1, 2, -3, 9}
        Dim allGood As Boolean = True

        For Each x In values
            If ((+x) <> x) Then allGood = False
            If ((-x) <> (0 - x)) Then allGood = False
            If (Math.Sign(x) > 0 AndAlso x <= 0) Then allGood = False
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn integer_arithmetic_matrix_type_boundary_checks() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim maxInt As Integer = Integer.MaxValue
        Dim minInt As Integer = Integer.MinValue

        Console.WriteLine(maxInt > minInt)
        Console.WriteLine(Integer.MinValue < 0)
        Console.WriteLine(CLng(maxInt) + 1 > maxInt)
        Console.WriteLine(CLng(maxInt) + 1 - 1 = CLng(maxInt))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True", "True"]);
}
