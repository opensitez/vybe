! vybe-test: fortran/arithmetic_extended/unary_minus_on_real_literal
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (abs((-3.5) - -3.5) > 1.0e-6) then
    print *, "FAIL: want [-3.5] got [", -3.5, "]"
    stop 1
end if
end program t
