! vybe-test: fortran/control_flow/named_do_cycle_outer
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
    integer :: i
    integer :: acc
    acc = 0

    outer: do i = 1, 4
        if (mod(i, 2) == 1) cycle outer
        acc = acc + i
    end do outer

        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((acc) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", acc, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
