! vybe-test: fortran/array_sections_extended/section_2d_1_to_2_comma_2_to_3_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 12, 23, 70 ]
integer :: a(3,4)
integer :: b(2,2)
integer :: i, j
do i = 1, 3
do j = 1, 4
a(i,j) = i * 10 + j
end do
end do
b = a(1:2, 2:3)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((b(1,1)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", b(1,1), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((b(2,2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", b(2,2), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((sum(b)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum(b), "]"
    stop 1
end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
