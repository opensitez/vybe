! vybe-test: fortran/array_reduction_extended/product_mask_select_three
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
logical :: m(5) = [.true., .true., .false., .true., .false.]
if ((product(a, mask=m)) /= 8) then
    print *, "FAIL: want [8] got [", product(a, mask=m), "]"
    stop 1
end if
end program t
