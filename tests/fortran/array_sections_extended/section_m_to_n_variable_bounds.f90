! vybe-test: fortran/array_sections_extended/section_m_to_n_variable_bounds
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
integer :: m, n
m = 2
n = 6
if ((size(a(m:n))) /= 5) then
    print *, "FAIL: want [5] got [", size(a(m:n)), "]"
    stop 1
end if
if ((sum(a(m:n))) /= 20) then
    print *, "FAIL: want [20] got [", sum(a(m:n)), "]"
    stop 1
end if
end program t
