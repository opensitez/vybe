! vybe-test: fortran/arrays_dim_mask/maxval_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_max(4)
    col_max = maxval(m, dim=1)
    print *, col_max(1)
end program test
