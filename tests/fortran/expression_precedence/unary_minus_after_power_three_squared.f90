! vybe-test: fortran/expression_precedence/unary_minus_after_power_three_squared
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((-3 ** 2) /= -9) then
    print *, "FAIL: want [-9] got [", -3 ** 2, "]"
    stop 1
end if
end program t
