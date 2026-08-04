! vybe-test: fortran/array_sections_extended/section_colon_n_first_three
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(7) = [(i, i = 1, 7)]
integer :: n
n = 3
if ((sum(a(:n))) /= 6) then
    print *, "FAIL: want [6] got [", sum(a(:n)), "]"
    stop 1
end if
end program t
