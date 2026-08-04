! vybe-test: fortran/do_loop_extended/do_concurrent_multiply_factor
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
integer :: a(6)
do concurrent (i = 1:6)
a(i) = i * 3
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((a(3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(3), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
