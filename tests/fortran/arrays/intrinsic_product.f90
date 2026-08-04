! vybe-test: fortran/arrays/intrinsic_product
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    print *, product(a)
end program test
