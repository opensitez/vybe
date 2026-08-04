! vybe-test: fortran/array_sections_extended/section_1_to_n_with_n_five
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(8) = [(i * 2, i = 1, 8)]
integer :: n
n = 5
if ((a(1:n)(5)) /= 10) then
    print *, "FAIL: want [10] got [", a(1:n)(5), "]"
    stop 1
end if
if ((sum(a(1:n))) /= 30) then
    print *, "FAIL: want [30] got [", sum(a(1:n)), "]"
    stop 1
end if
end program t
