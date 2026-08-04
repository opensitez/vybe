! vybe-test: fortran/arithmetic_extended/mixed_real_fractional_roundtrip
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (abs((7 / 2 + 0.5) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", 7 / 2 + 0.5, "]"
    stop 1
end if
end program t
