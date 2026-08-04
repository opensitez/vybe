! vybe-test: fortran/named_loops_extended/mixed_exit_and_cycle_named_outer
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 17 ]
integer :: i, j, s
s = 0
outer: do i = 1, 6
inner: do j = 1, 6
if (j == 1) cycle inner
if (j == 5) cycle outer
if (i == 5 .and. j == 3) exit outer
s = s + 1
end do inner
end do outer
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
