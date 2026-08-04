! vybe-test: fortran/ieee_intrinsics_extended/ieee_value_negative_zero_is_finite
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x
x = ieee_value(x, ieee_negative_zero)
if ((merge(1, 0, ieee_is_finite(x))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(x)), "]"
    stop 1
end if
end program t
