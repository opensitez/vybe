! vybe-test: fortran/array_sections_extended/assign_section_2_to_5_scalar
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [(i, i = 1, 8)]
a(2:5) = 7
if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 7) then
    print *, "FAIL: want [7] got [", a(2), "]"
    stop 1
end if
if ((a(5)) /= 7) then
    print *, "FAIL: want [7] got [", a(5), "]"
    stop 1
end if
if ((a(6)) /= 6) then
    print *, "FAIL: want [6] got [", a(6), "]"
    stop 1
end if
end program t
