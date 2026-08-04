! vybe-test: fortran/named_loops/named_do_named_while_runs_and_leaves_value
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
    integer :: n
    integer :: count
    n = 0
    count = 0
    counting: do while (n < 3)
        n = n + 1
        count = count + n
    end do counting
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((count) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", count, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
