! vybe-test: fortran/arithmetic_extended/power_right_associative_four_cubed_squared
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((4 ** (3 ** 2)) /= 262144) then
    print *, "FAIL: want [262144] got [", 4 ** (3 ** 2), "]"
    stop 1
end if
end program t
