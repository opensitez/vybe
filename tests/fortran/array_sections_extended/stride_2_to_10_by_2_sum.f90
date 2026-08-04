! vybe-test: fortran/array_sections_extended/stride_2_to_10_by_2_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
if ((sum(a(2:10:2))) /= 30) then
    print *, "FAIL: want [30] got [", sum(a(2:10:2)), "]"
    stop 1
end if
end program t
