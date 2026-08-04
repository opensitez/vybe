! vybe-test: fortran/expression_precedence/paren_negated_power_cube
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((-(2 ** 3)) /= -8) then
    print *, "FAIL: want [-8] got [", -(2 ** 3), "]"
    stop 1
end if
end program t
