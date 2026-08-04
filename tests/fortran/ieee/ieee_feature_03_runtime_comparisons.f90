! vybe-test: fortran/ieee/ieee_feature_03_runtime_comparisons
! origin: languages/fortran/tests/fortran/test_ieee.rs

program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
use, intrinsic :: ieee_arithmetic
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((ieee_rem(7.0, 3.0)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_rem(7.0, 3.0), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((ieee_is_finite(1.0)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_is_finite(1.0), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((ieee_is_normal(1.0)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_is_normal(1.0), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((ieee_unordered(1.0, 2.0)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_unordered(1.0, 2.0), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
