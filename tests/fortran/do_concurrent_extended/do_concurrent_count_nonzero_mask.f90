! vybe-test: fortran/do_concurrent_extended/do_concurrent_count_nonzero_mask
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
integer :: a(10), c, k
a = 0
do concurrent (i = 1:10, mod(i,2)==0)
a(i) = 1
end do
c = 0
do k = 1, 10
if (a(k) /= 0) c = c + 1
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
