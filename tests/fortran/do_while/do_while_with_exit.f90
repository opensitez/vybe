! vybe-test: fortran/do_while/do_while_with_exit
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 7 ]
    integer :: n = 0
    do while (.true.)
        n = n + 1
        if (n == 7) exit
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((n) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", n, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
