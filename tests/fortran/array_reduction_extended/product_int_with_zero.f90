! vybe-test: fortran/array_reduction_extended/product_int_with_zero
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(5) = [3, 5, 0, 7, 9]
if ((product(a)) /= 0) then
    print *, "FAIL: want [0] got [", product(a), "]"
    stop 1
end if
end program t
