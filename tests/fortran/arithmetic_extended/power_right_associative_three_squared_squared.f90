! vybe-test: fortran/arithmetic_extended/power_right_associative_three_squared_squared
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((3 ** (2 ** 2)) /= 81) then
    print *, "FAIL: want [81] got [", 3 ** (2 ** 2), "]"
    stop 1
end if
end program t
