! vybe-test: fortran/ieee_intrinsics_extended/huge_double_precision
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
integer, parameter :: dp = selected_real_kind(15)
if ((merge(1, 0, huge(1.0_dp) > 1.0e30)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, huge(1.0_dp) > 1.0e30), "]"
    stop 1
end if
end program t
