! vybe-test: fortran/array_reduction_extended/count_int_equal_to_two
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(6) = [1, 2, 2, 3, 2, 4]
if ((count(a == 2)) /= 3) then
    print *, "FAIL: want [3] got [", count(a == 2), "]"
    stop 1
end if
end program t
