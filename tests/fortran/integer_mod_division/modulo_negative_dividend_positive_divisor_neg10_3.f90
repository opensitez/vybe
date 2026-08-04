! vybe-test: fortran/integer_mod_division/modulo_negative_dividend_positive_divisor_neg10_3
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(-10, 3)) /= 2) then
    print *, "FAIL: want [2] got [", modulo(-10, 3), "]"
    stop 1
end if
end program t
