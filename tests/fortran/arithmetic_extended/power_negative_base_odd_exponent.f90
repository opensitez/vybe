! vybe-test: fortran/arithmetic_extended/power_negative_base_odd_exponent
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (((-2) ** 3) /= -8) then
    print *, "FAIL: want [-8] got [", (-2) ** 3, "]"
    stop 1
end if
end program t
