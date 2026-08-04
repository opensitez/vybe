! vybe-test: fortran/if_blocks/if_elseif_chain_missing_final_else
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(2) = [ "two", "done" ]
integer :: x = 2
if (x == 3) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim("three") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "three", "]"
    stop 1
end if
else if (x == 2) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim("two") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "two", "]"
    stop 1
end if
else if (x == 1) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim("one") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "one", "]"
    stop 1
end if
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if (trim("done") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "done", "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
