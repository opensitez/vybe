! vybe-test: fortran/fortran_control_flow_quality/control_flow_if_without_then_false_only_while
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_if_without_then_false_only_while
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer :: i
    integer :: total
    total = 0
    i = 1
    do while (i < 6)
        if (mod(i, 3) == 0) then
            total = total + i
        end if
        i = i + 1
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
end program control_flow_if_without_then_false_only_while
