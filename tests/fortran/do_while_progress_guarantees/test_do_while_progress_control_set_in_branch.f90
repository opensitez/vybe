! vybe-test: fortran/do_while_progress_guarantees/test_do_while_progress_control_set_in_branch
! origin: languages/fortran/tests/fortran/test_do_while_progress_guarantees.rs

    program test_do_while_progress_control_set_in_branch
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer :: n
    logical :: ok
    n = 0
    ok = .true.
    do while (ok)
        n = n + 1
        if (n >= 3) then
            ok = .false.
        else
            n = n + 1
        end if
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
end program test_do_while_progress_control_set_in_branch
