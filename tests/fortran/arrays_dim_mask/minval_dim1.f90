! vybe-test: fortran/arrays_dim_mask/minval_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_min(4)
    col_min = minval(m, dim=1)
    print *, col_min(2)
end program test
