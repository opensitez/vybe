! vybe-test: fortran/do_concurrent_extended/do_concurrent_max_via_sequential_scan
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 24 ]
integer :: a(8), mx, k
do concurrent (i = 1:8)
a(i) = i * 3
end do
mx = a(1)
do k = 2, 8
if (a(k) > mx) mx = a(k)
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((mx) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", mx, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
