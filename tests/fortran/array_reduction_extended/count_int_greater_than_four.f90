! vybe-test: fortran/array_reduction_extended/count_int_greater_than_four
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(7) = [1, 5, 3, 8, 2, 6, 4]
if ((count(a > 4)) /= 3) then
    print *, "FAIL: want [3] got [", count(a > 4), "]"
    stop 1
end if
end program t
