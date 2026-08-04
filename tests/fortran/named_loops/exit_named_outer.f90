! vybe-test: fortran/named_loops/exit_named_outer
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 3, 3 ]
    integer :: i, j
    outer: do i = 1, 5
        inner: do j = 1, 5
            if (i == 3 .and. j == 3) exit outer
        end do inner
    end do outer
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
    if ((j) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", j, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
