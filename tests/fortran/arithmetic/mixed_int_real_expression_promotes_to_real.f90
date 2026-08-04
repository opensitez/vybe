! vybe-test: fortran/arithmetic/mixed_int_real_expression_promotes_to_real
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((10 - 3 * 2.0 + 1) /= 5) then
    print *, "FAIL: want [5] got [", 10 - 3 * 2.0 + 1, "]"
    stop 1
end if
end program t
