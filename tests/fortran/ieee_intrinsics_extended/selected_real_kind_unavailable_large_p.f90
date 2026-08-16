! vybe-test: fortran/ieee_intrinsics_extended/selected_real_kind_unavailable_large_p
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((selected_real_kind(1000)) /= -1) then
    print *, "FAIL: want [-1] got [", selected_real_kind(1000), "]"
    stop 1
end if
end program t
