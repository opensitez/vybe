! vybe-test: fortran/ieee_intrinsics_extended/selected_real_kind_assign_to_param
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
integer, parameter :: sp = selected_real_kind(6, 37)
if ((sp) /= 4) then
    print *, "FAIL: want [4] got [", sp, "]"
    stop 1
end if
end program t
