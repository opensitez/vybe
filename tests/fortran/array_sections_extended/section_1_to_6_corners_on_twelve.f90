! vybe-test: fortran/array_sections_extended/section_1_to_6_corners_on_twelve
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(12) = [(i, i = 1, 12)]
if ((a(1:6)(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1:6)(1), "]"
    stop 1
end if
if ((a(1:6)(6)) /= 6) then
    print *, "FAIL: want [6] got [", a(1:6)(6), "]"
    stop 1
end if
if ((size(a(1:6))) /= 6) then
    print *, "FAIL: want [6] got [", size(a(1:6)), "]"
    stop 1
end if
end program t
