! vybe-test: fortran/implicit_none_scope_rules/test_implicit_none_allows_host_association_to_inner_without_redeclaration
! origin: languages/fortran/tests/fortran/test_implicit_none_scope_rules.rs

program test_implicit_none_host_assoc
    implicit none
    integer :: host_value = 42
    call print_host()

contains
    subroutine print_host()
        implicit none
        if ((host_value) /= 42) then
    print *, "FAIL: want [42] got [", host_value, "]"
    stop 1
end if
    end subroutine print_host
end program test_implicit_none_host_assoc
