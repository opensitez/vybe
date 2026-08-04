! vybe-test: fortran/modulo_dim_sign_extended/merge_real_do_loop_clamp_negatives
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 0, 20, 40 ]
real :: v(4)=[-1.5,2.0,-3.0,4.0]
real :: w(4)
integer :: i
do i=1,4
w(i) = merge(v(i), 0.0, v(i)>0.0)
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((nint(w(1)*10)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", nint(w(1)*10), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((nint(w(2)*10)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", nint(w(2)*10), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((nint(w(4)*10)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", nint(w(4)*10), "]"
    stop 1
end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
