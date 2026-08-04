! vybe-test: fortran/arithmetic/unary_minus_parenthesized_takes_precedence
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((-(2 + 3)) /= -5) then
    print *, "FAIL: want [-5] got [", -(2 + 3), "]"
    stop 1
end if
end program t
