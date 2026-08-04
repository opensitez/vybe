! vybe-test: fortran/array_sections_extended/assign_whole_colon_from_slice
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
integer :: b(3)
b = [9, 8, 7]
a(:) = 0
a(2:4) = b
if ((a(1)) /= 0) then
    print *, "FAIL: want [0] got [", a(1), "]"
    stop 1
end if
if ((a(3)) /= 8) then
    print *, "FAIL: want [8] got [", a(3), "]"
    stop 1
end if
if ((a(5)) /= 0) then
    print *, "FAIL: want [0] got [", a(5), "]"
    stop 1
end if
end program t
