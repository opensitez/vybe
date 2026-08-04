! vybe-test: fortran/integer_mod_division/mod_vs_modulo_negative_dividend_positive_divisor_neg7_3
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(-7, 3)) /= -1) then
    print *, "FAIL: want [-1] got [", mod(-7, 3), "]"
    stop 1
end if
if ((modulo(-7, 3)) /= 2) then
    print *, "FAIL: want [2] got [", modulo(-7, 3), "]"
    stop 1
end if
end program t
