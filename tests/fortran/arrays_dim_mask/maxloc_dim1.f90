! vybe-test: fortran/arrays_dim_mask/maxloc_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
    integer :: col_maxloc(3)
    col_maxloc = maxloc(m, dim=1)
    print *, col_maxloc(1)
end program test
