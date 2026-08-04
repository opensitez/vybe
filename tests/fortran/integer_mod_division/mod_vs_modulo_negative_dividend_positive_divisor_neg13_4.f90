! vybe-test: fortran/integer_mod_division/mod_vs_modulo_negative_dividend_positive_divisor_neg13_4
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(-13, 4)) /= -1) then
    print *, "FAIL: want [-1] got [", mod(-13, 4), "]"
    stop 1
end if
if ((modulo(-13, 4)) /= 3) then
    print *, "FAIL: want [3] got [", modulo(-13, 4), "]"
    stop 1
end if
end program t
