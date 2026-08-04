! vybe-test: fortran/integer_mod_division/modulo_both_negative_dividend_and_divisor_neg10_neg3
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(-10, -3)) /= -1) then
    print *, "FAIL: want [-1] got [", modulo(-10, -3), "]"
    stop 1
end if
end program t
