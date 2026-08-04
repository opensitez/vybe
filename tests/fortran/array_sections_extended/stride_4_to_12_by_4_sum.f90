! vybe-test: fortran/array_sections_extended/stride_4_to_12_by_4_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(12) = [(i, i = 1, 12)]
if ((sum(a(4:12:4))) /= 24) then
    print *, "FAIL: want [24] got [", sum(a(4:12:4)), "]"
    stop 1
end if
end program t
