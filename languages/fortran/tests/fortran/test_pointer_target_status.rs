use super::helpers::run_prints;

#[test]
fn test_pointer_target_status_checks_association() {
    let out = run_prints(
        r#"
program test_pointer_target_status
    integer, target :: storage
    integer, pointer :: p
    p => storage
    print *, associated(p)
    nullify(p)
    print *, associated(p)
end program test_pointer_target_status
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}
