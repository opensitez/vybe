! vybe-test: fortran/array_reduction_extended/sum_mask_even_positions
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
logical :: m(6) = [.false., .true., .false., .true., .false., .true.]
if ((sum(a, mask=m)) /= 12) then
    print *, "FAIL: want [12] got [", sum(a, mask=m), "]"
    stop 1
end if
end program t
