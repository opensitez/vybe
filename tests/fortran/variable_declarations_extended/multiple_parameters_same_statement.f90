! vybe-test: fortran/variable_declarations_extended/multiple_parameters_same_statement
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: a = 1, b = 2, c = 3
if ((a + b + c) /= 6) then
    print *, "FAIL: want [6] got [", a + b + c, "]"
    stop 1
end if
end program t
