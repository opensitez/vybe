! vybe-test: fortran/variable_declarations_extended/parameter_logical_false
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical, parameter :: nope = .false.
if ((nope) .neqv. .false.) then
    print *, "FAIL: want [false] got [", nope, "]"
    stop 1
end if
end program t
