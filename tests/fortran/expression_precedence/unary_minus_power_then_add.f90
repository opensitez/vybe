! vybe-test: fortran/expression_precedence/unary_minus_power_then_add
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((-2 ** 3 + 1) /= -7) then
    print *, "FAIL: want [-7] got [", -2 ** 3 + 1, "]"
    stop 1
end if
end program t
