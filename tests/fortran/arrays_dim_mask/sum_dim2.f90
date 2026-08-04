! vybe-test: fortran/arrays_dim_mask/sum_dim2
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: row_sums(3)
    row_sums = sum(m, dim=2)
    print *, row_sums(1)
end program test
