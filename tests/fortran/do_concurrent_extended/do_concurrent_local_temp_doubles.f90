! vybe-test: fortran/do_concurrent_extended/do_concurrent_local_temp_doubles
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
integer :: src(5), dst(5)
src = [1, 2, 3, 4, 5]
do concurrent (i = 1:5) local(tmp)
integer :: tmp
tmp = src(i) * 2
dst(i) = tmp
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((dst(3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", dst(3), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
