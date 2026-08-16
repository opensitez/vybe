! vybe-test: fortran/array_sections_extended/sum_section_n_to_end
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(6) = [2, 4, 6, 8, 10, 12]
integer :: n
n = 3
if ((sum(a(n:))) /= 36) then
    print *, "FAIL: want [36] got [", sum(a(n:)), "]"
    stop 1
end if
end program t
