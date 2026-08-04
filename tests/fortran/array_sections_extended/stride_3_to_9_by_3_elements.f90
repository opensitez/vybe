! vybe-test: fortran/array_sections_extended/stride_3_to_9_by_3_elements
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(9) = [(i, i = 1, 9)]
if ((a(3:9:3)(1)) /= 3) then
    print *, "FAIL: want [3] got [", a(3:9:3)(1), "]"
    stop 1
end if
if ((a(3:9:3)(2)) /= 6) then
    print *, "FAIL: want [6] got [", a(3:9:3)(2), "]"
    stop 1
end if
if ((a(3:9:3)(3)) /= 9) then
    print *, "FAIL: want [9] got [", a(3:9:3)(3), "]"
    stop 1
end if
end program t
