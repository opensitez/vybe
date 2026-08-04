! vybe-test: fortran/variable_declarations_extended/parameter_real_expression
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
real, parameter :: tau = 2.0 * 3.14159
if (abs((tau) - 6.28318) > 1.0e-6) then
    print *, "FAIL: want [6.28318] got [", tau, "]"
    stop 1
end if
end program t
