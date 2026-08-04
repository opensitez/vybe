! vybe-test: fortran/integer_mod_division/mod_positive_dividend_negative_divisor_11_neg4
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(11, -4)) /= 3) then
    print *, "FAIL: want [3] got [", mod(11, -4), "]"
    stop 1
end if
end program t
