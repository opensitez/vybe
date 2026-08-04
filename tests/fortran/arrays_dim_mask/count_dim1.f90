! vybe-test: fortran/arrays_dim_mask/count_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_count(4)
    col_count = count(m > 6, dim=1)
    print *, col_count(1)
    print *, col_count(3)
end program test
