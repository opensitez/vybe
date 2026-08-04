! vybe-test: fortran/array_sections_extended/section_2d_assign_submatrix
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 0, 0, 6 ]
integer :: a(3,3)
integer :: i, j
a = 0
do i = 1, 3
do j = 1, 3
a(i,j) = i + j
end do
end do
a(1:2, 2:3) = 0
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((a(1,2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(1,2), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((a(2,3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(2,3), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((a(3,3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a(3,3), "]"
    stop 1
end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
