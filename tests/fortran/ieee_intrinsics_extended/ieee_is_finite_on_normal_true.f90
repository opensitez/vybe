! vybe-test: fortran/ieee_intrinsics_extended/ieee_is_finite_on_normal_true
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
if ((merge(1, 0, ieee_is_finite(1.0))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(1.0)), "]"
    stop 1
end if
end program t
