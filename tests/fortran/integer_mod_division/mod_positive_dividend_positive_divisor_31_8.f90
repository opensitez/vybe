! vybe-test: fortran/integer_mod_division/mod_positive_dividend_positive_divisor_31_8
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(31, 8)) /= 7) then
    print *, "FAIL: want [7] got [", mod(31, 8), "]"
    stop 1
end if
end program t
