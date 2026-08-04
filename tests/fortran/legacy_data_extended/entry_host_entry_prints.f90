! vybe-test: fortran/legacy_data_extended/entry_host_entry_prints
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 100 ]
call worker()
contains
subroutine worker()
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((100) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 100, "]"
    stop 1
end if
return
entry alt_worker()
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((200) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 200, "]"
    stop 1
end if
end subroutine worker
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
