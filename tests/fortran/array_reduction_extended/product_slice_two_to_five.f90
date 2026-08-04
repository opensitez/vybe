! vybe-test: fortran/array_reduction_extended/product_slice_two_to_five
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(8) = [(i, i = 1, 8)]
if ((product(a(3:5))) /= 60) then
    print *, "FAIL: want [60] got [", product(a(3:5)), "]"
    stop 1
end if
end program t
