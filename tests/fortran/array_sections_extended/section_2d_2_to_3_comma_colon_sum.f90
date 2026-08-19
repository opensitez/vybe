! vybe-test: fortran/array_sections_extended/section_2d_2_to_3_comma_colon_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 27 ]
integer :: a(4,3)
integer :: i, j
do i = 1, 4
do j = 1, 3
a(i,j) = i + j
end do
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum(a(2:3, :))) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum(a(2:3, :)), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
