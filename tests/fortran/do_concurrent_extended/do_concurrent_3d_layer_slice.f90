! vybe-test: fortran/do_concurrent_extended/do_concurrent_3d_layer_slice
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
integer :: a(2,2,3)
a = 0
do concurrent (i = 1:2, j = 1:2, k = 1:3)
if (k == 2) a(i,j,k) = i + j
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((a(1,1,2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(1,1,2), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
