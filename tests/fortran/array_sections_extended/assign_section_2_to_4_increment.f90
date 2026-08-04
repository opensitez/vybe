! vybe-test: fortran/array_sections_extended/assign_section_2_to_4_increment
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
a(2:4) = a(2:4) + 10
if ((a(2)) /= 12) then
    print *, "FAIL: want [12] got [", a(2), "]"
    stop 1
end if
if ((a(4)) /= 14) then
    print *, "FAIL: want [14] got [", a(4), "]"
    stop 1
end if
if ((sum(a)) /= 45) then
    print *, "FAIL: want [45] got [", sum(a), "]"
    stop 1
end if
end program t
