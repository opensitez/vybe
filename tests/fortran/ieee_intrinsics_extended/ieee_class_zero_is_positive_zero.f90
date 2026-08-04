! vybe-test: fortran/ieee_intrinsics_extended/ieee_class_zero_is_positive_zero
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x = 0.0
if ((merge(1, 0, ieee_class(x) == ieee_positive_zero)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_class(x) == ieee_positive_zero), "]"
    stop 1
end if
end program t
