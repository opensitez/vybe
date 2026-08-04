! vybe-test: fortran/modulo_dim_sign_extended/do_real_modulo_wraps_angle_degrees
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
real :: a
integer :: i, c
c = 0
do i = 0, 359
a = real(i)
if (nint(modulo(a, 90.0)) == 0) c = c + 1
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((c) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", c, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
