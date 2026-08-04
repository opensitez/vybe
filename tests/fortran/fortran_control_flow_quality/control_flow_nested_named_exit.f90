! vybe-test: fortran/fortran_control_flow_quality/control_flow_nested_named_exit
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_nested_named_exit
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 17 ]
    integer :: outer_i, inner_i, total
    total = 0
    outer_loop: do outer_i = 1, 5
        do inner_i = 1, 5
            if (outer_i == 4 .and. inner_i == 3) exit outer_loop
            total = total + 1
        end do
    end do outer_loop
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program control_flow_nested_named_exit
