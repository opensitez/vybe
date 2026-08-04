! vybe-test: fortran/ieee_intrinsics_extended/ieee_value_copy_normal
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x = 7.0
if ((merge(1, 0, ieee_is_finite(ieee_value(x, ieee_positive_normal)))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(ieee_value(x, ieee_positive_normal))), "]"
    stop 1
end if
end program t
