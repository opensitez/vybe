! vybe-test: fortran/arithmetic_extended/power_before_addition
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((1 + 2 ** 4) /= 17) then
    print *, "FAIL: want [17] got [", 1 + 2 ** 4, "]"
    stop 1
end if
end program t
