! vybe-test: fortran/array_sections_extended/minval_section_3_to_7
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(9) = [9, 8, 1, 4, 2, 7, 3, 6, 5]
if ((minval(a(3:7))) /= 1) then
    print *, "FAIL: want [1] got [", minval(a(3:7)), "]"
    stop 1
end if
end program t
