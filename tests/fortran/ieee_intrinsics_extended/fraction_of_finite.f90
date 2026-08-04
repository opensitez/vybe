! vybe-test: fortran/ieee_intrinsics_extended/fraction_of_finite
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
use ieee_arithmetic
if ((merge(1, 0, ieee_is_finite(fraction(1.5)))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(fraction(1.5))), "]"
    stop 1
end if
end program t
