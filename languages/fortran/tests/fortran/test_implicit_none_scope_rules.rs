use super::helpers::run_prints;

#[test]
fn test_implicit_none_scope_rules_rejects_undeclared_but_scopes_locally() {
    let out = run_prints(
        r#"
program test_implicit_none_scope_rules
    implicit none
    integer :: outer
    outer = 5
    call inner(outer)

contains
    subroutine inner(v)
        implicit none
        integer, intent(in) :: v
        print *, v
    end subroutine
end program test_implicit_none_scope_rules
"#,
    );

    assert_eq!(out, vec!["5"]);
}
