! vybe-test: fortran/if_blocks/if_without_final_else_no_branch_taken
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "after" ]
if (1 == 0) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("should not print") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "should not print", "]"
    stop 1
end if
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("after") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "after", "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
