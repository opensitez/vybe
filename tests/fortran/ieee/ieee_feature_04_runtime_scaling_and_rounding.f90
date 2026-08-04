! vybe-test: fortran/ieee/ieee_feature_04_runtime_scaling_and_rounding
! origin: languages/fortran/tests/fortran/test_ieee.rs

program p
use, intrinsic :: ieee_arithmetic
if ((ieee_scalb(1.0, 2)) /= 4) then
    print *, "FAIL: want [4] got [", ieee_scalb(1.0, 2), "]"
    stop 1
end if
if ((ieee_rint(1.6)) /= 2) then
    print *, "FAIL: want [2] got [", ieee_rint(1.6), "]"
    stop 1
end if
if ((ieee_logb(8.0)) /= 3) then
    print *, "FAIL: want [3] got [", ieee_logb(8.0), "]"
    stop 1
end if
end program p
