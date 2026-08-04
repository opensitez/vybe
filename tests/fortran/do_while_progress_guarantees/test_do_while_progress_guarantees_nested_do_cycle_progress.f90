! vybe-test: fortran/do_while_progress_guarantees/test_do_while_progress_guarantees_nested_do_cycle_progress
! origin: languages/fortran/tests/fortran/test_do_while_progress_guarantees.rs

program test_do_while_progress_guarantees_nested_do_cycle_progress
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 4, 12 ]
    integer :: outer
    integer :: inner
    integer :: count
    outer = 0
    count = 0
    do while (outer < 4)
        outer = outer + 1
        inner = 0
        do while (inner < 4)
            inner = inner + 1
            if (inner == 2) cycle
            count = count + 1
        end do
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((outer) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", outer, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((count) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", count, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test_do_while_progress_guarantees_nested_do_cycle_progress
