! vybe-test: fortran/logical_eqv_neqv/if_eqv_true_false_prints_else_branch
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
integer :: vybe_check_i = 0
character(len=8) :: vybe_check_w(1) = [ "mismatch" ]
if (.true. .eqv. .false.) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("match") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "match", "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("mismatch") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "mismatch", "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
