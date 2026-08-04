! vybe-test: fortran/control/ctrl_select_type_20
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
class(*),allocatable::x
allocate(integer::x)
x = 1
select type(x)
 type is(integer)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((x) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", x, "]"
      stop 1
  end if
 class default
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((2) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 2, "]"
      stop 1
  end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
