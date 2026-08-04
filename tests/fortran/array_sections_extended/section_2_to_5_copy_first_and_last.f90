! vybe-test: fortran/array_sections_extended/section_2_to_5_copy_first_and_last
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [1,2,3,4,5,6,7,8]
integer :: b(4)
b = a(2:5)
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 5) then
    print *, "FAIL: want [5] got [", b(4), "]"
    stop 1
end if
end program t
