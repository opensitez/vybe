! vybe-test: fortran/implicit_none_scope_rules/test_implicit_none_keeps_inner_scope_variables_separate
! origin: languages/fortran/tests/fortran/test_implicit_none_scope_rules.rs

program test_implicit_none_scope_separation
    implicit none
    integer :: shared = 7
    integer :: result = 0
    call mutate(shared, result)
    if ((result) /= 8) then
    print *, "FAIL: want [8] got [", result, "]"
    stop 1
end if

contains
    subroutine mutate(inp, outp)
        integer, intent(in) :: inp
        integer, intent(out) :: outp
        integer :: local
        local = inp + 1
        outp = local
        if ((local) /= 8) then
    print *, "FAIL: want [8] got [", local, "]"
    stop 1
end if
    end subroutine mutate
end program test_implicit_none_scope_separation
