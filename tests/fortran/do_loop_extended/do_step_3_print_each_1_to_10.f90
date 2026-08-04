! vybe-test: fortran/do_loop_extended/do_step_3_print_each_1_to_10
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(4) = [ 1, 4, 7, 10 ]
integer :: i
do i = 1, 10, 3
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program t
