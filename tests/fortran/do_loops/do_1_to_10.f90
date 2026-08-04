! vybe-test: fortran/do_loops/do_1_to_10
! origin: languages/fortran/tests/fortran/test_do_loops.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 55 ]
integer :: i, s
s = 0
do i = 1, 10
s = s + i
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((s) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", s, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
