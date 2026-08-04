! vybe-test: fortran/control_flow_extended/select_type_integer_allocation_prints_int
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
character(len=3) :: vybe_check_w(1) = [ "int" ]
class(*), allocatable :: x
allocate(integer::x)
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
