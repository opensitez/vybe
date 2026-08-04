! vybe-test: fortran/integer_mod_division/modulo_both_negative_dividend_and_divisor_neg17_neg5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(-17, -5)) /= -2) then
    print *, "FAIL: want [-2] got [", modulo(-17, -5), "]"
    stop 1
end if
end program t
