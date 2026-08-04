! vybe-test: fortran/array_reduction_extended/count_gt_on_slice
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(7) = [1, 2, 3, 4, 5, 6, 7]
if ((count(a(2:6) > 3)) /= 3) then
    print *, "FAIL: want [3] got [", count(a(2:6) > 3), "]"
    stop 1
end if
end program t
