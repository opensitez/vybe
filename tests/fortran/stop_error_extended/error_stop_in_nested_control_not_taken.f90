! vybe-test: fortran/stop_error_extended/error_stop_in_nested_control_not_taken
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(2) = [ "ok", "done" ]
logical :: fail = .false.
logical :: inner = .false.
if (fail) then
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if (trim('bad') /= trim(vybe_check_w(vybe_check_i))) then
        print *, "FAIL at ", vybe_check_i, " got [", 'bad', "]"
        stop 1
    end if
else if (inner) then
    error stop 'nested bad'
else
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if (trim('ok') /= trim(vybe_check_w(vybe_check_i))) then
        print *, "FAIL at ", vybe_check_i, " got [", 'ok', "]"
        stop 1
    end if
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim('done') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'done', "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
