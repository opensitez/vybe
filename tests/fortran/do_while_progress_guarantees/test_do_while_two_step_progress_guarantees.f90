! vybe-test: fortran/do_while_progress_guarantees/test_do_while_two_step_progress_guarantees
! origin: languages/fortran/tests/fortran/test_do_while_progress_guarantees.rs

program test_do_while_progress_guarantees_two_step
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 10, 30 ]
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 9)
        i = i + 2
        total = total + i
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((i) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", i, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test_do_while_progress_guarantees_two_step
