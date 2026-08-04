! vybe-test: fortran/ieee_intrinsics_extended/epsilon_zero_arg_positive
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((merge(1, 0, epsilon(0.0) > 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, epsilon(0.0) > 0.0), "]"
    stop 1
end if
end program t
