! vybe-test: fortran/implicit_none_scope_rules/test_implicit_none_scope_rules_rejects_undeclared_but_scopes_locally
! origin: languages/fortran/tests/fortran/test_implicit_none_scope_rules.rs

program test_implicit_none_scope_rules
    implicit none
    integer :: outer
    outer = 5
    call inner(outer)

contains
    subroutine inner(v)
        implicit none
        integer, intent(in) :: v
        if ((v) /= 5) then
    print *, "FAIL: want [5] got [", v, "]"
    stop 1
end if
    end subroutine
end program test_implicit_none_scope_rules
