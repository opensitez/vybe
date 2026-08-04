! vybe-test: fortran/array_sections_extended/assign_section_from_other_section
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(6) = [6, 5, 4, 3, 2, 1]
integer :: b(3)
b = a(2:4)
a(4:6) = b
if ((a(4)) /= 5) then
    print *, "FAIL: want [5] got [", a(4), "]"
    stop 1
end if
if ((a(5)) /= 4) then
    print *, "FAIL: want [4] got [", a(5), "]"
    stop 1
end if
if ((a(6)) /= 3) then
    print *, "FAIL: want [3] got [", a(6), "]"
    stop 1
end if
end program t
