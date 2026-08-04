! vybe-test: fortran/arithmetic_extended/power_times_multiplier
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((2 ** 3 * 4) /= 32) then
    print *, "FAIL: want [32] got [", 2 ** 3 * 4, "]"
    stop 1
end if
end program t
