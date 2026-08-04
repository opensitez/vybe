! vybe-test: fortran/full_programs/power_of_two_table
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(4) = [ 1, 2, 4, 8 ]
integer :: i, p
p = 1
do i = 0, 3
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if ((p) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", p, "]"
    stop 1
end if
p = p * 2
end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program t
