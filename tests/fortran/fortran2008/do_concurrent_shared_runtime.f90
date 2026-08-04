! vybe-test: fortran/fortran2008/do_concurrent_shared_runtime
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 3, 15 ]
    integer :: a(5)
    integer :: factor
    integer :: i
    factor = 3
    do concurrent (i = 1:5) shared(factor)
        a(i) = i * factor
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
