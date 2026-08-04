! vybe-test: fortran/do_while/do_while_with_logical_not_condition
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    logical :: keep_running
    integer :: n
    keep_running = .true.
    n = 0
    do while (.not. .not. keep_running)
        n = n + 1
        if (n == 3) keep_running = .false.
        if (n > 5) exit
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
