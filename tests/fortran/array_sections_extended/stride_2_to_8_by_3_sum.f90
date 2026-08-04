! vybe-test: fortran/array_sections_extended/stride_2_to_8_by_3_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
if ((sum(a(2:8:3))) /= 15) then
    print *, "FAIL: want [15] got [", sum(a(2:8:3)), "]"
    stop 1
end if
end program t
