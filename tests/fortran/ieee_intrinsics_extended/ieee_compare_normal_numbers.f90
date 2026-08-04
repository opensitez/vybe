! vybe-test: fortran/ieee_intrinsics_extended/ieee_compare_normal_numbers
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
real :: x = 2.0, y = 3.0
if ((merge(1, 0, ieee_is_finite(x + y))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(x + y)), "]"
    stop 1
end if
end program t
