! vybe-test: fortran/arrays_dim_mask/sum_dim_mask_combined
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(4,4) = reshape([(i, i=1,16)],[4,4])
    logical :: mask(4,4)
    integer :: row_sums(4)
    mask = m > 8
    row_sums = sum(m, dim=2, mask=mask)
    print *, row_sums(1)
end program test
