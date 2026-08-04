! vybe-test: fortran/do_while_progress_guarantees/test_do_while_progress_guarantees_nested_progress_after_cycle
! origin: languages/fortran/tests/fortran/test_do_while_progress_guarantees.rs

program test_do_while_progress_guarantees_nested_progress_after_cycle
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 3, 3 ]
    integer :: outer
    integer :: inner
    integer :: total
    outer = 0
    total = 0
    do while (outer < 3)
        outer = outer + 1
        inner = 0
        do while (inner < 4)
            inner = inner + 1
            if (inner == 2) cycle
            if (inner == 3) exit
            total = total + 1
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
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test_do_while_progress_guarantees_nested_progress_after_cycle
