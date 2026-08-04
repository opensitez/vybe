! vybe-test: fortran/do_loops/do_exit
! origin: languages/fortran/tests/fortran/test_do_loops.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 1, 2, 3 ]
integer :: i
do i = 1, 100
if (i > 3) exit
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
end do
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
