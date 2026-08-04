! vybe-test: fortran/fortran_control_flow_quality/control_flow_logical_precedence
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_logical_precedence
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
    logical :: a, b, c
    a = .true.
    b = .false.
    c = .true.
    if (a .and. b .or. c) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program control_flow_logical_precedence
