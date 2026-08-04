! vybe-test: fortran/array_sections_extended/section_4_to_9_six_elements
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(12) = [(i, i = 1, 12)]
if ((size(a(4:9))) /= 6) then
    print *, "FAIL: want [6] got [", size(a(4:9)), "]"
    stop 1
end if
if ((sum(a(4:9))) /= 39) then
    print *, "FAIL: want [39] got [", sum(a(4:9)), "]"
    stop 1
end if
end program t
