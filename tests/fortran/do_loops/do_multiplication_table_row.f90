! vybe-test: fortran/do_loops/do_multiplication_table_row
! origin: languages/fortran/tests/fortran/test_do_loops.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(5) = [ 3, 6, 9, 12, 15 ]
integer :: i
do i = 1, 5
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 5) then
    print *, "FAIL: more than 5 line(s)"
    stop 1
end if
if ((3 * i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 3 * i, "]"
    stop 1
end if
end do
if (vybe_check_i /= 5) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 5"
    stop 1
end if
end program t
