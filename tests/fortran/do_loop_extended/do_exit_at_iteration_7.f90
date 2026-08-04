! vybe-test: fortran/do_loop_extended/do_exit_at_iteration_7
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(6) = [ 1, 2, 3, 4, 5, 6 ]
integer :: i
do i = 1, 20
if (i == 7) exit
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 6) then
    print *, "FAIL: more than 6 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
end do
if (vybe_check_i /= 6) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 6"
    stop 1
end if
end program t
