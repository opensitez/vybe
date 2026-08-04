! vybe-test: fortran/control/ctrl_exit_nested_21
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 7 ]
integer::outer_i, inner_i, total
total = 0
outer: do outer_i = 1, 4
  do inner_i = 1, 3
    if (outer_i == 3 .and. inner_i == 2) exit outer
    total = total + 1
  end do
end do outer
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
