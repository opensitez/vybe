! vybe-test: fortran/control_flow_extended/select_case_range_branches
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "three" ]
integer :: day = 15
select case (day)
case (1:7)
  vybe_check_i = vybe_check_i + 1
 if (vybe_check_i > 1) then
     print *, "FAIL: more than 1 line(s)"
     stop 1
 end if
 if (trim('week') /= trim(vybe_check_w(vybe_check_i))) then
     print *, "FAIL at ", vybe_check_i, " got [", 'week', "]"
     stop 1
 end if
case (8:14)
  vybe_check_i = vybe_check_i + 1
 if (vybe_check_i > 1) then
     print *, "FAIL: more than 1 line(s)"
     stop 1
 end if
 if (trim('half') /= trim(vybe_check_w(vybe_check_i))) then
     print *, "FAIL at ", vybe_check_i, " got [", 'half', "]"
     stop 1
 end if
case (15:21)
  vybe_check_i = vybe_check_i + 1
 if (vybe_check_i > 1) then
     print *, "FAIL: more than 1 line(s)"
     stop 1
 end if
 if (trim('three') /= trim(vybe_check_w(vybe_check_i))) then
     print *, "FAIL at ", vybe_check_i, " got [", 'three', "]"
     stop 1
 end if
case default
  vybe_check_i = vybe_check_i + 1
 if (vybe_check_i > 1) then
     print *, "FAIL: more than 1 line(s)"
     stop 1
 end if
 if (trim('other') /= trim(vybe_check_w(vybe_check_i))) then
     print *, "FAIL at ", vybe_check_i, " got [", 'other', "]"
     stop 1
 end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
