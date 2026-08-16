! vybe-test: fortran/named_loops_extended/exit_outer_leaves_mid_and_inner
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 20 ]
integer :: i, j, k, c
c = 0
outer: do i = 1, 4
mid: do j = 1, 4
inner: do k = 1, 4
if (i == 2 .and. j == 2 .and. k == 1) exit outer
c = c + 1
end do inner
end do mid
end do outer
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
