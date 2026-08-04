! vybe-test: fortran/integer_mod_division/mod_zero_result_exact_multiple_36_9
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((mod(36, 9)) /= 0) then
    print *, "FAIL: want [0] got [", mod(36, 9), "]"
    stop 1
end if
end program t
