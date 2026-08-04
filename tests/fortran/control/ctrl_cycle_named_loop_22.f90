! vybe-test: fortran/control/ctrl_cycle_named_loop_22
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
integer::i, j, total
total = 0
row: do i = 1, 2
  col: do j = 1, 5
    if (mod(j, 2) == 0) cycle row
    total = total + 1
  end do col
end do row
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((total) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", total, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
