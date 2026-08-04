! vybe-test: fortran/array_sections_extended/section_n_to_end_runtime_n_three
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
integer :: n
n = 3
if ((a(n:)(1)) /= 3) then
    print *, "FAIL: want [3] got [", a(n:)(1), "]"
    stop 1
end if
if ((a(n:)(3)) /= 5) then
    print *, "FAIL: want [5] got [", a(n:)(3), "]"
    stop 1
end if
end program t
