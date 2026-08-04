! vybe-test: fortran/fortran_control_flow_quality/control_flow_nested_if_chain_with_elseif
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_nested_if_chain
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer :: x
    x = 17
    if (x > 20) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    else if (x > 10) then
        if (mod(x, 2) == 0) then
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((2) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 2, "]"
                stop 1
            end if
        else
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((3) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 3, "]"
                stop 1
            end if
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((4) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 4, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program control_flow_nested_if_chain
