! vybe-test: fortran/ieee_intrinsics_extended/selected_real_kind_p15_r500
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((selected_real_kind(15, 500)) /= 16) then
    print *, "FAIL: want [16] got [", selected_real_kind(15, 500), "]"
    stop 1
end if
end program t
