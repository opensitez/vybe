! vybe-test: fortran/named_loops/cycle_preserves_state
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
    integer :: i, j, sum_j
    sum_j = 0
    outer: do i = 1, 4
        inner: do j = 1, 4
            if (mod(j, 2) == 0) cycle outer
            sum_j = sum_j + j
        end do inner
    end do outer
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((sum_j) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum_j, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
