! vybe-test: fortran/array_reduction_extended/product_int_negatives_pair
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(4) = [-2, 3, -4, 5]
if ((product(a)) /= 120) then
    print *, "FAIL: want [120] got [", product(a), "]"
    stop 1
end if
end program t
