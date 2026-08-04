! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_opened_false_after_close
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
logical :: opened
open(51, status='scratch')
close(51)
inquire(unit=51, opened=opened)
if (opened) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((1) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
