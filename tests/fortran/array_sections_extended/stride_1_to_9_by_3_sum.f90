! vybe-test: fortran/array_sections_extended/stride_1_to_9_by_3_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(9) = [(i, i = 1, 9)]
if ((sum(a(1:9:3))) /= 12) then
    print *, "FAIL: want [12] got [", sum(a(1:9:3)), "]"
    stop 1
end if
end program t
