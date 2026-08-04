! vybe-test: fortran/array_reduction_extended/product_int_powers_of_two
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(4) = [2, 4, 8, 16]
if ((product(a)) /= 1024) then
    print *, "FAIL: want [1024] got [", product(a), "]"
    stop 1
end if
end program t
