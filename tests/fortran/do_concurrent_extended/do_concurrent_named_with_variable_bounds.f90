! vybe-test: fortran/do_concurrent_extended/do_concurrent_named_with_variable_bounds
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 8 ]
integer :: a(7), n
n = 7
fill: do concurrent (i = 1:n)
a(i) = 2 * i
end do fill
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((a(4)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(4), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
