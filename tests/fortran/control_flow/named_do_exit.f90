! vybe-test: fortran/control_flow/named_do_exit
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
    integer :: i
    integer :: j
    outer: do i = 1, 5
        inner: do j = 1, 5
            if (i == 2 .and. j == 3) exit outer
        end do inner
    end do outer
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((i) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", i, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
