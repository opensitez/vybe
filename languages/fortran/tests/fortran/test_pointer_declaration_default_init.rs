use super::helpers::run_prints;

#[test]
fn test_pointer_declaration_default_init_is_associated_state() {
    let out = run_prints(
        r#"
program test_pointer_declaration_default_init
    integer, target :: storage
    integer, pointer :: p => null()
    if (associated(p)) then
        print *, 0
    else
        print *, 1
    end if
    p => storage
    storage = 9
    print *, p
end program test_pointer_declaration_default_init
"#,
    );

    assert_eq!(out, vec!["1", "9"]);
}
