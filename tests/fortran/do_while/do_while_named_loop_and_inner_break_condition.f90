! vybe-test: fortran/do_while/do_while_named_loop_and_inner_break_condition
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 8 ]
    integer :: i, j, s
    s = 0
    i = 0
    outer: do while (i < 5)
        i = i + 1
        j = 0
        do while (j < 5)
            j = j + 1
            if (j == 3) exit
            s = s + 1
        end do
    end do outer
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((s) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", s, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
