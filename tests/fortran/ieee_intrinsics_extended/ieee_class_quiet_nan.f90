! vybe-test: fortran/ieee_intrinsics_extended/ieee_class_quiet_nan
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x
x = ieee_value(x, ieee_quiet_nan)
if ((merge(1, 0, ieee_is_nan(x) .or. ieee_is_finite(x) .eqv. .false.)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_nan(x) .or. ieee_is_finite(x) .eqv. .false.), "]"
    stop 1
end if
end program t
