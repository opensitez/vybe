! vybe-test: fortran/ieee_intrinsics_extended/selected_real_kind_p4_r10
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((selected_real_kind(4, 10)) /= 4) then
    print *, "FAIL: want [4] got [", selected_real_kind(4, 10), "]"
    stop 1
end if
end program t
