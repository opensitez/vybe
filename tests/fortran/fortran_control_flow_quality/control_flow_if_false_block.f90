! vybe-test: fortran/fortran_control_flow_quality/control_flow_if_false_block
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_if_false_block
    integer :: value
    if (.false.) then
        value = 1
    else
        value = 2
    end if
    if ((value) /= 2) then
    print *, "FAIL: want [2] got [", value, "]"
    stop 1
end if
end program control_flow_if_false_block
