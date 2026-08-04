! vybe-test: fortran/integer_mod_division/mod_zero_result_exact_multiple_20_5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(20, 5)) /= 0) then
    print *, "FAIL: want [0] got [", mod(20, 5), "]"
    stop 1
end if
end program t
