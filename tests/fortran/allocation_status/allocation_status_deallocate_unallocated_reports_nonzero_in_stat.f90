! vybe-test: fortran/allocation_status/allocation_status_deallocate_unallocated_reports_nonzero_in_stat
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program t
integer :: vybe_check_i = 0
character(len=3) :: vybe_check_w(1) = [ "bad" ]
integer, allocatable :: a(:)
integer :: st
deallocate(a, stat=st)
if (st == 0) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('ok') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'ok', "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('bad') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'bad', "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
