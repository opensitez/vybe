! vybe-test: fortran/array_sections_extended/sum_section_2_to_5_seven_elements
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(7) = [3, 1, 9, 1, 5, 8, 2]
if ((sum(a(2:5))) /= 16) then
    print *, "FAIL: want [16] got [", sum(a(2:5)), "]"
    stop 1
end if
end program t
