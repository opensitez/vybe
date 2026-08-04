! vybe-test: fortran/arrays_dim_mask/product_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
    integer :: col_prod(3)
    col_prod = product(m, dim=1)
    print *, col_prod(1)
end program test
