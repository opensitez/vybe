! vybe-test: fortran/if_blocks/if_logical_parentheses_and_not
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: vybe_check_i = 0
logical :: vybe_check_w(1) = [ .false. ]
if ((1 > 0) .and. .not. (2 < 3)) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (("true") .neqv. vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", "true", "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (("false") .neqv. vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", "false", "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
