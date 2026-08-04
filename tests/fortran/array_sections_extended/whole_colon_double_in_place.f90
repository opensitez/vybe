! vybe-test: fortran/array_sections_extended/whole_colon_double_in_place
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(3) = [5, 10, 15]
a(:) = a(:) * 2
if ((a(1)) /= 10) then
    print *, "FAIL: want [10] got [", a(1), "]"
    stop 1
end if
if ((a(3)) /= 30) then
    print *, "FAIL: want [30] got [", a(3), "]"
    stop 1
end if
if ((sum(a)) /= 60) then
    print *, "FAIL: want [60] got [", sum(a), "]"
    stop 1
end if
end program t
