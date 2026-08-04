! vybe-test: fortran/do_concurrent_extended/do_concurrent_sum_diagonal_2d
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 15 ]
integer :: m(5,5), s, k
m = 0
do concurrent (i = 1:5, j = 1:5)
if (i == j) m(i,j) = i
end do
s = 0
do k = 1, 5
s = s + m(k,k)
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
