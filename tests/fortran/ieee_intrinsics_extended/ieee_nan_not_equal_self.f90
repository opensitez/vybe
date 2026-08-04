! vybe-test: fortran/ieee_intrinsics_extended/ieee_nan_not_equal_self
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x
x = ieee_value(x, ieee_quiet_nan)
if ((merge(1, 0, x == x)) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, x == x), "]"
    stop 1
end if
end program t
