! vybe-test: fortran/stop_error_extended/guarded_stop_after_warning_print
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
integer :: vybe_check_i = 0
character(len=7) :: vybe_check_w(1) = [ "warning" ]
integer :: n = 2
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('warning') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'warning', "]"
    stop 1
end if
if (n > 1) stop 1
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('tail') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'tail', "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
