! vybe-test: fortran/integer_mod_division/integer_division_truncates_toward_zero_neg17_div_neg5
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((-17 / -5) /= 3) then
    print *, "FAIL: want [3] got [", -17 / -5, "]"
    stop 1
end if
end program t
