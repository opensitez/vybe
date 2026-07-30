use super::helpers::run_prints;

#[test]
fn test_io_format_edit_descriptor_matrix_mix_int_and_real() {
    let out = run_prints(
        r#"
program test_io_format_edit_descriptor_matrix
    integer :: n
    real :: x
    n = 12
    x = 4.5
    print '(I4,1X,F4.1)', n, x
end program test_io_format_edit_descriptor_matrix
"#,
    );

    assert_eq!(out, vec!["0012 4.5"]);
}

#[test]
fn test_io_format_edit_descriptor_matrix_string_with_repeat() {
    let out = run_prints(
        r#"
program test_io_format_edit_descriptor_matrix
    print '(A,1X,A,1X,A)', "a", "b", "c"
end program test_io_format_edit_descriptor_matrix
"#,
    );

    assert_eq!(out, vec!["a b c"]);
}

#[test]
fn test_io_format_edit_descriptor_matrix_float_and_int_mixed() {
    let out = run_prints(
        r#"
program test_io_format_edit_descriptor_matrix
    print '(I0,\",\",E8.2)', 4, 1.0
end program test_io_format_edit_descriptor_matrix
"#,
    );

    assert_eq!(out, vec!["4,0.10E+01"]);
}
