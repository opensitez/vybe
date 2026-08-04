! vybe-test: fortran/arithmetic_extended/integer_power_real_exponent
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((2 ** 3.0) /= 8) then
    print *, "FAIL: want [8] got [", 2 ** 3.0, "]"
    stop 1
end if
end program t
