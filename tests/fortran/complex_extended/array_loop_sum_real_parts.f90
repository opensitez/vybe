! vybe-test: fortran/complex_extended/array_loop_sum_real_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
complex :: x(3)
integer :: i
real :: s
x(1) = cmplx(1.0, 0.0)
x(2) = cmplx(2.0, 0.0)
x(3) = cmplx(3.0, 0.0)
s = 0.0
do i = 1, 3
  s = s + real(x(i))
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((nint(s)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", nint(s), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
