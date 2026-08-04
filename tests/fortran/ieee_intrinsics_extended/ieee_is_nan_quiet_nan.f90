! vybe-test: fortran/ieee_intrinsics_extended/ieee_is_nan_quiet_nan
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x
x = ieee_value(x, ieee_quiet_nan)
if ((merge(1, 0, ieee_is_nan(x))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_nan(x)), "]"
    stop 1
end if
if ((merge(1, 0, ieee_is_finite(x))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, ieee_is_finite(x)), "]"
    stop 1
end if
end program t
