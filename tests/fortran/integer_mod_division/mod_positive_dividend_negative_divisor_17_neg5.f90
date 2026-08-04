! vybe-test: fortran/integer_mod_division/mod_positive_dividend_negative_divisor_17_neg5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(17, -5)) /= 2) then
    print *, "FAIL: want [2] got [", mod(17, -5), "]"
    stop 1
end if
end program t
