! vybe-test: fortran/do_while/do_while_named_outer_exit_from_inner
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 2, 4 ]
    integer :: outer, inner, total
    outer = 0
    total = 0
    pump: do while (outer < 10)
        outer = outer + 1
        inner = 0
        do while (inner < 5)
            inner = inner + 1
            if (outer == 2 .and. inner == 3) exit pump
            total = total + 1
        end do
    end do pump
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
end program test
