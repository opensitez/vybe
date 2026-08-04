! vybe-test: fortran/integer_mod_division/mod_positive_dividend_positive_divisor_14_6
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(14, 6)) /= 2) then
    print *, "FAIL: want [2] got [", mod(14, 6), "]"
    stop 1
end if
end program t
