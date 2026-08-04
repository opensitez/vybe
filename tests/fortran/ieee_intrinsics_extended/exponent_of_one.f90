! vybe-test: fortran/ieee_intrinsics_extended/exponent_of_one
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((exponent(1.0)) /= 1) then
    print *, "FAIL: want [1] got [", exponent(1.0), "]"
    stop 1
end if
end program t
