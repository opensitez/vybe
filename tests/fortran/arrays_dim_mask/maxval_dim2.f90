! vybe-test: fortran/arrays_dim_mask/maxval_dim2
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: row_max(3)
    row_max = maxval(m, dim=2)
    print *, row_max(1)
end program test
