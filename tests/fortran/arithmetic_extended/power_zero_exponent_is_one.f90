! vybe-test: fortran/arithmetic_extended/power_zero_exponent_is_one
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((5 ** 0) /= 1) then
    print *, "FAIL: want [1] got [", 5 ** 0, "]"
    stop 1
end if
end program t
