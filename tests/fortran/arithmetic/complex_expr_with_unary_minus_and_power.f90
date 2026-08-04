! vybe-test: fortran/arithmetic/complex_expr_with_unary_minus_and_power
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((-2 ** 3 + 1) /= -7) then
    print *, "FAIL: want [-7] got [", -2 ** 3 + 1, "]"
    stop 1
end if
end program t
