! vybe-test: fortran/control/ctrl_select_case_14
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
integer::x=2
select case(x)
 case(1)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((1) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
      stop 1
  end if
 case(2)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((2) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 2, "]"
      stop 1
  end if
 case default
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((3) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 3, "]"
      stop 1
  end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
