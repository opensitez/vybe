! vybe-test: fortran/logical_eqv_neqv/if_neqv_nested_inside_eqv_picks_else
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "else" ]
if ((.true. .neqv. .false.) .eqv. .false.) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("then") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "then", "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("else") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "else", "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
