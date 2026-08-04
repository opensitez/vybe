! vybe-test: fortran/array_reduction_extended/product_int_one_to_five
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(5) = [(i, i = 1, 5)]
if ((product(a)) /= 120) then
    print *, "FAIL: want [120] got [", product(a), "]"
    stop 1
end if
end program t
