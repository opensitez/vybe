! vybe-test: fortran/named_loops/three_level_exit_middle
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 16 ]
    integer :: i, j, k, count
    count = 0
    outer: do i = 1, 4
        mid: do j = 1, 4
            inner: do k = 1, 4
                if (j == 2 .and. k == 2) exit mid
                count = count + 1
            end do inner
        end do mid
    end do outer
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((count) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", count, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
