! vybe-test: fortran/do_concurrent_extended/do_concurrent_locality_runtime
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 2, 10 ]
    integer :: a(5), b(5)
    integer :: i
    b = [1, 2, 3, 4, 5]
    do concurrent (i = 1:5) local(tmp)
        integer :: tmp
        tmp = b(i) * 2
        a(i) = tmp
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((a(1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", a(1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((a(5)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", a(5), "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
