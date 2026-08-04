! vybe-test: fortran/control_flow_extended/select_type_character_defaults_to_class_default
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "other" ]
class(*), allocatable :: x
allocate(character(len=4)::x)
select type(x)
 type is(integer)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim('int') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'int', "]"
      stop 1
  end if
 type is(real)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim('real') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'real', "]"
      stop 1
  end if
 class default
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
