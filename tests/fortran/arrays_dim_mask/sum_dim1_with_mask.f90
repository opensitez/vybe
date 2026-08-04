! vybe-test: fortran/arrays_dim_mask/sum_dim1_with_mask
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
    logical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
    integer :: col_sums(3)
    col_sums = sum(m, dim=1, mask=mask)
    print *, col_sums(1)
end program test
