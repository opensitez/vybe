! vybe-test: fortran/control_flow_extended/labeled_continue_after_print_in_do
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 1, 2, 4 ]
integer :: i
do i = 1, 4
if (i == 3) goto 200
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
200 continue
end do
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
