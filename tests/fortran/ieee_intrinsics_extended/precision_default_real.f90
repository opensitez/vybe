! vybe-test: fortran/ieee_intrinsics_extended/precision_default_real
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((precision(1.0)) /= 6) then
    print *, "FAIL: want [6] got [", precision(1.0), "]"
    stop 1
end if
end program t
