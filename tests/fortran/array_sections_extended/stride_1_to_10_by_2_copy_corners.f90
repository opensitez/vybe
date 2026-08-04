! vybe-test: fortran/array_sections_extended/stride_1_to_10_by_2_copy_corners
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
integer :: b(5)
b = a(1:10:2)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(5)) /= 9) then
    print *, "FAIL: want [9] got [", b(5), "]"
    stop 1
end if
if ((size(b)) /= 5) then
    print *, "FAIL: want [5] got [", size(b), "]"
    stop 1
end if
end program t
