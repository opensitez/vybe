! vybe-test: fortran/control_flow/do_loop_zero_iterations_with_positive_step_on_descending_bounds
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
    integer :: i
    integer :: s
    s = 0
    do i = 10, 1, 1
        s = s + 1
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((s) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", s, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
