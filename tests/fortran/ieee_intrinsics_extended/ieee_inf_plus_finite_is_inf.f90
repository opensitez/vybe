! vybe-test: fortran/ieee_intrinsics_extended/ieee_inf_plus_finite_is_inf
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x, y
x = ieee_value(x, ieee_positive_inf)
y = 1.0
if ((merge(1, 0, ieee_is_finite(x + y))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, ieee_is_finite(x + y)), "]"
    stop 1
end if
end program t
