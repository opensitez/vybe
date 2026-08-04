! vybe-test: fortran/implicit_none_scope_rules/test_implicit_none_not_inherited_into_internal_subprogram
! origin: languages/fortran/tests/fortran/test_implicit_none_scope_rules.rs

program test_implicit_none_scope_inheritance
    implicit none
    call show_implicit_behavior()

contains
    subroutine show_implicit_behavior()
        integer :: explicit_x
        implicit integer (a-z)
        explicit_x = 3
        y = explicit_x
        if ((y) /= 3) then
    print *, "FAIL: want [3] got [", y, "]"
    stop 1
end if
    end subroutine show_implicit_behavior
end program test_implicit_none_scope_inheritance
