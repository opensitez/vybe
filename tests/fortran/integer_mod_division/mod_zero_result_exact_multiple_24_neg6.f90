! vybe-test: fortran/integer_mod_division/mod_zero_result_exact_multiple_24_neg6
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(24, -6)) /= 0) then
    print *, "FAIL: want [0] got [", mod(24, -6), "]"
    stop 1
end if
end program t
