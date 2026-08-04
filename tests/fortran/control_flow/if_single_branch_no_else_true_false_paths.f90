! vybe-test: fortran/control_flow/if_single_branch_no_else_true_false_paths
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
character(len=11) :: vybe_check_w(1) = [ "nonpositive" ]
    integer :: x
    x = 0
    if (x > 0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('positive') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'positive', "]"
            stop 1
        end if
    end if
    if (x <= 0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('nonpositive') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'nonpositive', "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
