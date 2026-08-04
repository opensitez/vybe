! vybe-test: fortran/integer_mod_division/modulo_positive_dividend_positive_divisor_23_7
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(23, 7)) /= 2) then
    print *, "FAIL: want [2] got [", modulo(23, 7), "]"
    stop 1
end if
end program t
