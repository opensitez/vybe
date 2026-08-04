! vybe-test: fortran/select_case/select_case_in_loop
! origin: languages/fortran/tests/fortran/test_select_case.rs

program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(4) = [ "odd", "even", "odd", "even" ]
integer :: i
do i = 1, 4
 select case (mod(i, 2))
 case (0)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 4) then
      print *, "FAIL: more than 4 line(s)"
      stop 1
  end if
  if (trim('even') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'even', "]"
      stop 1
  end if
 case (1)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 4) then
      print *, "FAIL: more than 4 line(s)"
      stop 1
  end if
  if (trim('odd') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'odd', "]"
      stop 1
  end if
 end select
end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program t
