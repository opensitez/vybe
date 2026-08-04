! vybe-test: fortran/array_sections_extended/whole_colon_scalar_assign_zero
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(4) = [9, 8, 7, 6]
a(:) = 0
if ((a(1)) /= 0) then
    print *, "FAIL: want [0] got [", a(1), "]"
    stop 1
end if
if ((a(4)) /= 0) then
    print *, "FAIL: want [0] got [", a(4), "]"
    stop 1
end if
if ((sum(a)) /= 0) then
    print *, "FAIL: want [0] got [", sum(a), "]"
    stop 1
end if
end program t
