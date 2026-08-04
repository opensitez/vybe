! vybe-test: fortran/arrays_dim_mask/sum_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_sums(4)
    col_sums = sum(m, dim=1)
    print *, col_sums(1)
end program test
