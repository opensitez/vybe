! vybe-test: fortran/ieee_intrinsics_extended/range_default_real
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((range(1.0)) /= 37) then
    print *, "FAIL: want [37] got [", range(1.0), "]"
    stop 1
end if
end program t
