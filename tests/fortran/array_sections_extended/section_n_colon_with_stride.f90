! vybe-test: fortran/array_sections_extended/section_n_colon_with_stride
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
integer :: n
n = 2
if ((sum(a(n:10:2))) /= 30) then
    print *, "FAIL: want [30] got [", sum(a(n:10:2)), "]"
    stop 1
end if
end program t
