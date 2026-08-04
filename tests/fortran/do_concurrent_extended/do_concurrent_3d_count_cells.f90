! vybe-test: fortran/do_concurrent_extended/do_concurrent_3d_count_cells
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 27 ]
integer :: a(3,3,3), c, ii, jj, kk
a = 1
c = 0
do ii = 1, 3
do jj = 1, 3
do kk = 1, 3
c = c + a(ii,jj,kk)
end do
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
