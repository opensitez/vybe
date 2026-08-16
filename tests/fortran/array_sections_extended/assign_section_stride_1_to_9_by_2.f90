! vybe-test: fortran/array_sections_extended/assign_section_stride_1_to_9_by_2
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(9) = [(i, i = 1, 9)]
a(1:9:2) = [11, 22, 33, 44, 55]
if ((a(1)) /= 11) then
    print *, "FAIL: want [11] got [", a(1), "]"
    stop 1
end if
if ((a(3)) /= 22) then
    print *, "FAIL: want [22] got [", a(3), "]"
    stop 1
end if
if ((a(9)) /= 55) then
    print *, "FAIL: want [55] got [", a(9), "]"
    stop 1
end if
if ((sum(a)) /= 185) then
    print *, "FAIL: want [185] got [", sum(a), "]"
    stop 1
end if
end program t
