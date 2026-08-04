! vybe-test: fortran/named_loops/named_do_outer_cycle_skips_to_next_outer_iteration
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer :: i, j, s
    s = 0
    outer: do i = 1, 3
        inner: do j = 1, 4
            if (j == 2) cycle outer
            s = s + 1
        end do inner
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
