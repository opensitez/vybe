! vybe-test: fortran/arithmetic_extended/power_two_to_eight
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((2 ** 8) /= 256) then
    print *, "FAIL: want [256] got [", 2 ** 8, "]"
    stop 1
end if
end program t
