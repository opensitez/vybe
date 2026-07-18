use super::helpers::run_prints;

#[test]
fn test_implicit_typing_recovery_paths_preserves_expression_types() {
    let out = run_prints(
        r#"
program test_implicit_typing_recovery_paths
    implicit none
    integer :: whole
    real :: fractional
    whole = 7 + 3
    fractional = real(whole) / 2.0
    print *, whole
    print *, nint(fractional)
end program test_implicit_typing_recovery_paths
"#,
    );

    assert_eq!(out, vec!["10", "5"]);
}
