! vybe-test: fortran/associatable_variable_lifetimes/test_associatable_variable_lifetimes_preserves_outer_binding
! origin: languages/fortran/tests/fortran/test_associatable_variable_lifetimes.rs

program test_associatable_variable_lifetimes
    implicit none
    integer :: base
    integer :: result
    base = 4
    associate(value => base)
        value = value + 1
        result = value
    end associate
    if ((result) /= 5) then
    print *, "FAIL: want [5] got [", result, "]"
    stop 1
end if
    if ((base) /= 5) then
    print *, "FAIL: want [5] got [", base, "]"
    stop 1
end if
end program test_associatable_variable_lifetimes
