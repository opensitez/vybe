! vybe-test: fortran/integer_mod_division/modulo_zero_result_exact_multiple_20_5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((modulo(20, 5)) /= 0) then
    print *, "FAIL: want [0] got [", modulo(20, 5), "]"
    stop 1
end if
end program t
