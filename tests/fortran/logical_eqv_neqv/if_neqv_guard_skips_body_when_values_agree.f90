! vybe-test: fortran/logical_eqv_neqv/if_neqv_guard_skips_body_when_values_agree
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "done" ]
if (.true. .neqv. .true.) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("run") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "run", "]"
    stop 1
end if
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("done") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "done", "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
