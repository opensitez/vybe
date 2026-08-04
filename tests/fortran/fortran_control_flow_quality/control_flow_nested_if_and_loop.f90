! vybe-test: fortran/fortran_control_flow_quality/control_flow_nested_if_and_loop
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_nested_if_and_loop
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 12 ]
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (mod(i, 2) == 0) then
            if (i <= 6) total = total + i
        end if
    end do
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
end program control_flow_nested_if_and_loop
