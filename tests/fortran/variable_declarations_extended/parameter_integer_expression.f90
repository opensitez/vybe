! vybe-test: fortran/variable_declarations_extended/parameter_integer_expression
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: n = 2 + 3
if ((n) /= 5) then
    print *, "FAIL: want [5] got [", n, "]"
    stop 1
end if
end program t
