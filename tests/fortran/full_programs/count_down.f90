! vybe-test: fortran/full_programs/count_down
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(5) = [ 5, 4, 3, 2, 1 ]
integer :: i
i = 5
do while (i > 0)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 5) then
    print *, "FAIL: more than 5 line(s)"
    stop 1
end if
if ((i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i, "]"
    stop 1
end if
i = i - 1
end do
if (vybe_check_i /= 5) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 5"
    stop 1
end if
end program t
