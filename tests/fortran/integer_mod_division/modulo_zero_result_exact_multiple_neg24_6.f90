! vybe-test: fortran/integer_mod_division/modulo_zero_result_exact_multiple_neg24_6
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(-24, 6)) /= 0) then
    print *, "FAIL: want [0] got [", modulo(-24, 6), "]"
    stop 1
end if
end program t
