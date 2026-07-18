use super::helpers::run_prints;

#[test]
fn test_declaration_statement_ordering_preserves_use_before_declare() {
    let out = run_prints(
        r#"
program test_declaration_statement_ordering
    integer :: value
    real :: ratio
    integer, parameter :: offset = 1

    value = 10
    ratio = real(value + offset) / 2.0

    print *, value
    print *, offset
    print *, nint(ratio)
end program test_declaration_statement_ordering
"#,
    );

    assert_eq!(out, vec!["10", "1", "5"]);
}
