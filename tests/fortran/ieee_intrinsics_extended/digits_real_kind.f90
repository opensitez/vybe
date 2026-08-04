! vybe-test: fortran/ieee_intrinsics_extended/digits_real_kind
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((digits(1.0)) /= 24) then
    print *, "FAIL: want [24] got [", digits(1.0), "]"
    stop 1
end if
end program t
