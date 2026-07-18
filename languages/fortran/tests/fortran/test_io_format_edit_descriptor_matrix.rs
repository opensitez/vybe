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
