! vybe-test: fortran/fortran2008/do_concurrent_mask
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
    integer :: a(10)
    a = 0
    do concurrent (i = 1:10, mod(i, 2) == 0)
        a(i) = i
    end do
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
end program test
