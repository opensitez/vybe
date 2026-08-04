! vybe-test: fortran/control_flow_extended/nested_cycle_inner_skips_j_equals_two
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
integer :: i, j, c
c = 0
do i = 1, 3
do j = 1, 4
if (j == 2) cycle
c = c + 1
end do
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((c) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", c, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
