! vybe-test: fortran/fortran_control_flow_quality/control_flow_mixed_assignment_flow
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_mixed_assignment_flow
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 22 ]
    integer :: x
    integer :: y
    x = 5
    y = 0
    if (x > 3) then
        y = x * 2
    else
        y = x + 1
    end if
    do while (y < 20)
        y = y + 3
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((y) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", y, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program control_flow_mixed_assignment_flow
