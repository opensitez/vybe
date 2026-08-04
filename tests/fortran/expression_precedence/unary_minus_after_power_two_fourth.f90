! vybe-test: fortran/expression_precedence/unary_minus_after_power_two_fourth
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((-2 ** 4) /= -16) then
    print *, "FAIL: want [-16] got [", -2 ** 4, "]"
    stop 1
end if
end program t
