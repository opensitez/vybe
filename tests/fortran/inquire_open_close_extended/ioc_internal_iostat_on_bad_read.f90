! vybe-test: fortran/inquire_open_close_extended/ioc_internal_iostat_on_bad_read
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
character(len=5) :: buf
integer :: n, ios
buf = 'abc'
read(buf, '(I0)', iostat=ios) n
if (ios /= 0) then
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
