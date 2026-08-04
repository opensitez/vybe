! vybe-test: fortran/arithmetic_extended/real_division_fifteen_over_six
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (abs((15.0 / 6.0) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", 15.0 / 6.0, "]"
    stop 1
end if
end program t
