! vybe-test: fortran/arithmetic_extended/real_division_seven_over_two
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (abs((7.0 / 2.0) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", 7.0 / 2.0, "]"
    stop 1
end if
end program t
