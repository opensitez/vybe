! vybe-test: fortran/if_construct_extended/if_compound_not_and_both_required
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "pass" ]
if (.not. (3 > 5 .and. 1 > 2)) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("pass") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "pass", "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("fail") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "fail", "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
