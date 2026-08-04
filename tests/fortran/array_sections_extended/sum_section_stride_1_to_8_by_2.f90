! vybe-test: fortran/array_sections_extended/sum_section_stride_1_to_8_by_2
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [(i, i = 1, 8)]
if ((sum(a(1:8:2))) /= 16) then
    print *, "FAIL: want [16] got [", sum(a(1:8:2)), "]"
    stop 1
end if
end program t
