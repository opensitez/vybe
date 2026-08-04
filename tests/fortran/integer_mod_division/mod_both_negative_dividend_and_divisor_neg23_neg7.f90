! vybe-test: fortran/integer_mod_division/mod_both_negative_dividend_and_divisor_neg23_neg7
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(-23, -7)) /= -2) then
    print *, "FAIL: want [-2] got [", mod(-23, -7), "]"
    stop 1
end if
end program t
