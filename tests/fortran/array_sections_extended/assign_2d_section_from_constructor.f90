! vybe-test: fortran/array_sections_extended/assign_2d_section_from_constructor
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(2,3)
a = 0
a(1:2, 2:3) = reshape([5, 6, 7, 8], [2, 2])
if ((a(1,2)) /= 5) then
    print *, "FAIL: want [5] got [", a(1,2), "]"
    stop 1
end if
if ((a(2,3)) /= 8) then
    print *, "FAIL: want [8] got [", a(2,3), "]"
    stop 1
end if
if ((sum(a)) /= 26) then
    print *, "FAIL: want [26] got [", sum(a), "]"
    stop 1
end if
end program t
