! vybe-test: fortran/do_while_progress_guarantees/test_do_while_progress_guarantees_condition_flip_break
! origin: languages/fortran/tests/fortran/test_do_while_progress_guarantees.rs

program test_do_while_progress_guarantees_condition_flip_break
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
    integer :: n
    logical :: active
    n = 0
    active = .true.
    do while (active)
        n = n + 1
        if (n == 6) active = .false.
        if (mod(n, 2) == 1) cycle
        if (n > 8) active = .false.
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
end program test_do_while_progress_guarantees_condition_flip_break
