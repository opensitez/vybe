! vybe-test: fortran/variable_declarations_extended/dimension_1d_integer_fill
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 30 ]
integer, dimension(4) :: a
integer :: i
do i = 1, 4
  a(i) = i * 10
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((a(3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(3), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
