! vybe-test: fortran/array_sections_extended/section_3_to_7_sum_ten_elements
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
if ((sum(a(3:7))) /= 25) then
    print *, "FAIL: want [25] got [", sum(a(3:7)), "]"
    stop 1
end if
end program t
