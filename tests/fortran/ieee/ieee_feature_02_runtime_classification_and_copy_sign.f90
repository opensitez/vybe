! vybe-test: fortran/ieee/ieee_feature_02_runtime_classification_and_copy_sign
! origin: languages/fortran/tests/fortran/test_ieee.rs

program p
integer :: vybe_check_i = 0
character(len=6) :: vybe_check_w(2) = [ "normal", "normal" ]
use, intrinsic :: ieee_arithmetic
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim(ieee_class(1.0)) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_class(1.0), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim(ieee_class(-1.0)) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_class(-1.0), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim(ieee_copy_sign(1.0, -2.0)) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", ieee_copy_sign(1.0, -2.0), "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program p
