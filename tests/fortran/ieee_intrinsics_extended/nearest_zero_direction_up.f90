! vybe-test: fortran/ieee_intrinsics_extended/nearest_zero_direction_up
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((merge(1, 0, nearest(0.0, 1.0) > 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, nearest(0.0, 1.0) > 0.0), "]"
    stop 1
end if
end program t
