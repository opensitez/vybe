! vybe-test: fortran/do_while/do_while_and_condition_with_parentheses
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 2, 2 ]
    integer :: a
    integer :: b
    a = 0
    b = 0
    do while (a < 5 .and. (b < 2))
        a = a + 1
        b = b + 1
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((a) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", a, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((b) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", b, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
