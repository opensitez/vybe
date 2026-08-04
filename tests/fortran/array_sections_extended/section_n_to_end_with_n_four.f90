! vybe-test: fortran/array_sections_extended/section_n_to_end_with_n_four
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(6) = [10, 20, 30, 40, 50, 60]
integer :: n
n = 4
if ((sum(a(n:))) /= 150) then
    print *, "FAIL: want [150] got [", sum(a(n:)), "]"
    stop 1
end if
if ((size(a(n:))) /= 3) then
    print *, "FAIL: want [3] got [", size(a(n:)), "]"
    stop 1
end if
end program t
