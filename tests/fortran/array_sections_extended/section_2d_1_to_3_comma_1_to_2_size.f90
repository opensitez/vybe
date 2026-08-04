! vybe-test: fortran/array_sections_extended/section_2d_1_to_3_comma_1_to_2_size
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(4,5)
integer :: b(3,2)
a = 1
b = a(1:3, 1:2)
if ((size(b)) /= 6) then
    print *, "FAIL: want [6] got [", size(b), "]"
    stop 1
end if
if ((sum(b)) /= 6) then
    print *, "FAIL: want [6] got [", sum(b), "]"
    stop 1
end if
end program t
