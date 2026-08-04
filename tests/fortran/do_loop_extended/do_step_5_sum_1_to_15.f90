! vybe-test: fortran/do_loop_extended/do_step_5_sum_1_to_15
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 18 ]
integer :: i, s
s = 0
do i = 1, 15, 5
s = s + i
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
end program t
