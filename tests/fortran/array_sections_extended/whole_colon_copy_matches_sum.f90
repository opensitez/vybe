! vybe-test: fortran/array_sections_extended/whole_colon_copy_matches_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(5) = [2, 4, 6, 8, 10]
integer :: b(5)
b = a(:)
if ((sum(b)) /= 30) then
    print *, "FAIL: want [30] got [", sum(b), "]"
    stop 1
end if
if ((b(3)) /= 6) then
    print *, "FAIL: want [6] got [", b(3), "]"
    stop 1
end if
end program t
