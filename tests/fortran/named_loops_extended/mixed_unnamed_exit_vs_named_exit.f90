! vybe-test: fortran/named_loops_extended/mixed_unnamed_exit_vs_named_exit
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
integer :: i, j, c
c = 0
named: do i = 1, 4
do j = 1, 4
if (j == 2) exit
c = c + 1
if (j == 3) exit named
end do
end do named
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
