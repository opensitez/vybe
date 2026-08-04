! vybe-test: fortran/integer_mod_division/integer_division_truncates_toward_zero_neg7_div_neg2
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
if ((-7 / -2) /= 3) then
    print *, "FAIL: want [3] got [", -7 / -2, "]"
    stop 1
end if
end program t
