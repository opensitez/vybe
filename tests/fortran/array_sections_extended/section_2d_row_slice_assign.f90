! vybe-test: fortran/array_sections_extended/section_2d_row_slice_assign
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(2,4)
a = 0
a(1, :) = [1, 2, 3, 4]
if ((a(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1,1), "]"
    stop 1
end if
if ((a(1,4)) /= 4) then
    print *, "FAIL: want [4] got [", a(1,4), "]"
    stop 1
end if
if ((sum(a(1,:))) /= 10) then
    print *, "FAIL: want [10] got [", sum(a(1,:)), "]"
    stop 1
end if
end program t
