! vybe-test: fortran/array_sections_extended/whole_colon_assign_from_constructor
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(4)
a(:) = [11, 22, 33, 44]
if ((a(2)) /= 22) then
    print *, "FAIL: want [22] got [", a(2), "]"
    stop 1
end if
if ((sum(a(:))) /= 110) then
    print *, "FAIL: want [110] got [", sum(a(:)), "]"
    stop 1
end if
end program t
