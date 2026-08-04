! vybe-test: fortran/arithmetic_extended/power_negative_base_even_exponent
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (((-2) ** 4) /= 16) then
    print *, "FAIL: want [16] got [", (-2) ** 4, "]"
    stop 1
end if
end program t
