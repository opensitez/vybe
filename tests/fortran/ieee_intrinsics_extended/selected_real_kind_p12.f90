! vybe-test: fortran/ieee_intrinsics_extended/selected_real_kind_p12
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((selected_real_kind(12)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(12), "]"
    stop 1
end if
end program t
