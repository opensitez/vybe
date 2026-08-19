! vybe-test: fortran/control_flow_extended/loop_exit_on_named_outer_target_hit
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 3, 2 ]
integer :: i, j
outer_loop: do i = 1, 4
inner_loop: do j = 1, 6
if (i == 3 .and. j == 2) exit outer_loop
end do inner_loop
end do outer_loop
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((j) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", j, "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
