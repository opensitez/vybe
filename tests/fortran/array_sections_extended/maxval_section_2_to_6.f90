! vybe-test: fortran/array_sections_extended/maxval_section_2_to_6
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [3, 1, 9, 1, 5, 8, 2, 7]
if ((maxval(a(2:6))) /= 9) then
    print *, "FAIL: want [9] got [", maxval(a(2:6)), "]"
    stop 1
end if
end program t
