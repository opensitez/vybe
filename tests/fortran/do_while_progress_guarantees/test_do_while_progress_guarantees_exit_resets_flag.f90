! vybe-test: fortran/do_while_progress_guarantees/test_do_while_progress_guarantees_exit_resets_flag
! origin: languages/fortran/tests/fortran/test_do_while_progress_guarantees.rs

program test_do_while_progress_guarantees_exit_resets_flag
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer :: n
    logical :: keep
    n = 0
    keep = .true.
    do while (keep)
        n = n + 1
        if (n == 3) keep = .false.
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
end program test_do_while_progress_guarantees_exit_resets_flag
