! vybe-test: fortran/control/ctrl_do_named_while_named_cycle_28
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
integer :: i, c
i = 0
c = 0
spin: do while (i < 6)
    i = i + 1
    if (mod(i, 2) == 0) cycle spin
    c = c + 1
end do spin
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((c) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", c, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
