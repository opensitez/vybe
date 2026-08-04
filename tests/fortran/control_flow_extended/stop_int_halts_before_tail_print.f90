! vybe-test: fortran/control_flow_extended/stop_int_halts_before_tail_print
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
character(len=6) :: vybe_check_w(1) = [ "before" ]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('before') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'before', "]"
    stop 1
end if
stop 0
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('after') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'after', "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
