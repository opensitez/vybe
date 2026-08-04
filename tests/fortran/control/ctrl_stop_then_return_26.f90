! vybe-test: fortran/control/ctrl_stop_then_return_26
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
character(len=6) :: vybe_check_w(1) = [ "before" ]
integer :: i
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('before') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'before', "]"
    stop 1
end if
i = 1
if (i == 1) then
    return
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim(i) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
