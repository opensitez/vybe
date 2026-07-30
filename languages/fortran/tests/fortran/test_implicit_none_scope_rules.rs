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

#[test]
fn test_implicit_none_allows_host_association_to_inner_without_redeclaration() {
    let out = run_prints(
        r#"
program test_implicit_none_host_assoc
    implicit none
    integer :: host_value = 42
    call print_host()

contains
    subroutine print_host()
        implicit none
        print *, host_value
    end subroutine print_host
end program test_implicit_none_host_assoc
"#,
    );

    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_implicit_none_not_inherited_into_internal_subprogram() {
    let out = run_prints(
        r#"
program test_implicit_none_scope_inheritance
    implicit none
    call show_implicit_behavior()

contains
    subroutine show_implicit_behavior()
        integer :: explicit_x
        implicit integer (a-z)
        explicit_x = 3
        y = explicit_x
        print *, y
    end subroutine show_implicit_behavior
end program test_implicit_none_scope_inheritance
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_implicit_none_keeps_inner_scope_variables_separate() {
    let out = run_prints(
        r#"
program test_implicit_none_scope_separation
    implicit none
    integer :: shared = 7
    integer :: result = 0
    call mutate(shared, result)
    print *, result

contains
    subroutine mutate(inp, outp)
        integer, intent(in) :: inp
        integer, intent(out) :: outp
        integer :: local
        local = inp + 1
        outp = local
        print *, local
    end subroutine mutate
end program test_implicit_none_scope_separation
"#,
    );

    assert_eq!(out, vec!["8", "8"]);
}
