! vybe-test: fortran/control_flow/if_elseif_chain
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
character(len=1) :: vybe_check_w(1) = [ "C" ]
    integer :: score
    score = 75
    if (score >= 90) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("A") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "A", "]"
            stop 1
        end if
    else if (score >= 80) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("B") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "B", "]"
            stop 1
        end if
    else if (score >= 70) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("C") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "C", "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("F") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "F", "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
