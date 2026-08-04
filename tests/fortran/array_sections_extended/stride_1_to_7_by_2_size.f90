! vybe-test: fortran/array_sections_extended/stride_1_to_7_by_2_size
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(7) = [(i, i = 1, 7)]
if ((size(a(1:7:2))) /= 4) then
    print *, "FAIL: want [4] got [", size(a(1:7:2)), "]"
    stop 1
end if
if ((sum(a(1:7:2))) /= 16) then
    print *, "FAIL: want [16] got [", sum(a(1:7:2)), "]"
    stop 1
end if
end program t
