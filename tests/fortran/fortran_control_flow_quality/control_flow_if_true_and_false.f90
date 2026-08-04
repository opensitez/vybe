! vybe-test: fortran/fortran_control_flow_quality/control_flow_if_true_and_false
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_if_true_and_false
    integer :: value
    if (.true.) then
        value = 1
    else
        value = 2
    end if
    if ((value) /= 1) then
    print *, "FAIL: want [1] got [", value, "]"
    stop 1
end if
end program control_flow_if_true_and_false
