! vybe-test: fortran/ieee_intrinsics_extended/ieee_is_nan_on_normal_false
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
if ((merge(1, 0, ieee_is_nan(1.0))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, ieee_is_nan(1.0)), "]"
    stop 1
end if
end program t
