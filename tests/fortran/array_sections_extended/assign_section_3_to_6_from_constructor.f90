! vybe-test: fortran/array_sections_extended/assign_section_3_to_6_from_constructor
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [(i, i = 1, 8)]
a(3:6) = [100, 200, 300, 400]
if ((a(2)) /= 2) then
    print *, "FAIL: want [2] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 100) then
    print *, "FAIL: want [100] got [", a(3), "]"
    stop 1
end if
if ((a(6)) /= 400) then
    print *, "FAIL: want [400] got [", a(6), "]"
    stop 1
end if
if ((a(7)) /= 7) then
    print *, "FAIL: want [7] got [", a(7), "]"
    stop 1
end if
end program t
