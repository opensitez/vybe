! vybe-test: fortran/arithmetic_extended/unary_minus_after_parentheses_with_real_multiplication
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((-(2.0 * 3.0) + 1) /= -5) then
    print *, "FAIL: want [-5] got [", -(2.0 * 3.0) + 1, "]"
    stop 1
end if
end program t
