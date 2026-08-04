! vybe-test: fortran/array_sections_extended/sum_whole_colon_section
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(7) = [(i, i = 1, 7)]
if ((sum(a(:))) /= 28) then
    print *, "FAIL: want [28] got [", sum(a(:)), "]"
    stop 1
end if
end program t
