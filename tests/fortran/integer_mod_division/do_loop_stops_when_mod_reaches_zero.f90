! vybe-test: fortran/integer_mod_division/do_loop_stops_when_mod_reaches_zero
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
integer :: i, c
c = 0
do i = 1, 100
if (mod(i, 7) == 0) c = c + 1
if (c == 5) exit
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
