! vybe-test: fortran/array_sections_extended/section_2_to_5_sum_on_eight_elements
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [1,2,3,4,5,6,7,8]
if ((sum(a(2:5))) /= 14) then
    print *, "FAIL: want [14] got [", sum(a(2:5)), "]"
    stop 1
end if
end program t
