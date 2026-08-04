! vybe-test: fortran/do_loop_extended/do_var_bounds_expr_step
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 12 ]
integer :: i, n, s
integer :: step
n = 7
step = 3
s = 0
do i = 1, n, step
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
