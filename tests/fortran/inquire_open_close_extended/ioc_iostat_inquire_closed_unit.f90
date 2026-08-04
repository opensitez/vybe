! vybe-test: fortran/inquire_open_close_extended/ioc_iostat_inquire_closed_unit
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 0, 0 ]
integer :: ios
logical :: opened
open(67, status='scratch')
close(67)
inquire(unit=67, opened=opened, iostat=ios)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((ios) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", ios, "]"
    stop 1
end if
if (opened) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((1) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
