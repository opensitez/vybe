! vybe-test: fortran/do_concurrent_extended/do_concurrent_2d_diagonal
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
integer :: m(4,4)
m = 0
do concurrent (i = 1:4, j = 1:4)
if (i == j) m(i,j) = i
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((m(3,3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", m(3,3), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
