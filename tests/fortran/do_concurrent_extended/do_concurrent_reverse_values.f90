! vybe-test: fortran/do_concurrent_extended/do_concurrent_reverse_values
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
integer :: a(5), b(5)
do concurrent (i = 1:5)
a(i) = i
end do
do concurrent (i = 1:5)
b(i) = a(6 - i)
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((b(1)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", b(1), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
