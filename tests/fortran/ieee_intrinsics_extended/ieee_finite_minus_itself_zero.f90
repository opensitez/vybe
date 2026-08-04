! vybe-test: fortran/ieee_intrinsics_extended/ieee_finite_minus_itself_zero
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x = 5.0
if ((merge(1, 0, ieee_is_finite(x - x))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(x - x)), "]"
    stop 1
end if
end program t
