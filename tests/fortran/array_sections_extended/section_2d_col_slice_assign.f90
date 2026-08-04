! vybe-test: fortran/array_sections_extended/section_2d_col_slice_assign
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(4,2)
a = 0
a(:, 2) = [5, 6, 7, 8]
if ((a(2,2)) /= 6) then
    print *, "FAIL: want [6] got [", a(2,2), "]"
    stop 1
end if
if ((a(4,2)) /= 8) then
    print *, "FAIL: want [8] got [", a(4,2), "]"
    stop 1
end if
if ((sum(a(:,2))) /= 26) then
    print *, "FAIL: want [26] got [", sum(a(:,2)), "]"
    stop 1
end if
end program t
