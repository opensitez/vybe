! vybe-test: fortran/integer_mod_division/modulo_negative_dividend_positive_divisor_neg17_5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(-17, 5)) /= 3) then
    print *, "FAIL: want [3] got [", modulo(-17, 5), "]"
    stop 1
end if
end program t
