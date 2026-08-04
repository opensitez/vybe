! vybe-test: fortran/named_loops_extended/mixed_named_mid_unnamed_inner_outer
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
integer :: i, j, k, c
c = 0
outer: do i = 1, 2
do j = 1, 2
mid: do k = 1, 3
if (k == 2) exit mid
c = c + 1
end do mid
end do
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
