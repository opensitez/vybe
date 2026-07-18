use super::helpers::run_prints;

#[test]
fn test_declaration_specifier_conflicts_parameter_and_storage() {
    let out = run_prints(
        r#"
program test_declaration_specifier_conflicts
    integer, parameter :: a = 7
    integer, target :: b
    integer, pointer :: p
    b = a
    p => b
    print *, a
    print *, p
end program test_declaration_specifier_conflicts
"#,
    );

    assert_eq!(out, vec!["7", "7"]);
}
