! vybe-test: fortran/do_while/do_while_with_cycle
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 25 ]
    integer :: i = 0, s = 0
    do while (i < 10)
        i = i + 1
        if (mod(i, 2) == 0) cycle
        s = s + i
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((s) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", s, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
