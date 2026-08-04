! vybe-test: fortran/do_loop_extended/do_exit_after_five_prints
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(5) = [ 1, 2, 3, 4, 5 ]
integer :: i, c
c = 0
do i = 1, 100
if (c == 5) exit
c = c + 1
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 5) then
    print *, "FAIL: more than 5 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
end do
if (vybe_check_i /= 5) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 5"
    stop 1
end if
end program t
