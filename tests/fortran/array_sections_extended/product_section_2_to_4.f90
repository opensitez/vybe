! vybe-test: fortran/array_sections_extended/product_section_2_to_4
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
if ((product(a(2:4))) /= 24) then
    print *, "FAIL: want [24] got [", product(a(2:4)), "]"
    stop 1
end if
end program t
