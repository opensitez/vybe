! vybe-test: fortran/random_number_extended/random_do_loop_ten_values_bounded
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
real :: r
integer :: i, ok
ok = 1
do i = 1, 10
  call random_number(r)
  if (r < 0.0 .or. r >= 1.0) ok = 0
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((ok) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", ok, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
