! vybe-test: fortran/array_reduction_extended/product_real_three_halves
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(3) = [2.0, 2.5, 2.0]
if ((product(a)) /= 10) then
    print *, "FAIL: want [10] got [", product(a), "]"
    stop 1
end if
end program t
