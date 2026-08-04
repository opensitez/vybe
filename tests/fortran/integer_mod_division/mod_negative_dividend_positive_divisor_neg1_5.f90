! vybe-test: fortran/integer_mod_division/mod_negative_dividend_positive_divisor_neg1_5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(-1, 5)) /= -1) then
    print *, "FAIL: want [-1] got [", mod(-1, 5), "]"
    stop 1
end if
end program t
