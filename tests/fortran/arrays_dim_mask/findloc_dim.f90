! vybe-test: fortran/arrays_dim_mask/findloc_dim
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,3) = reshape([1,2,1,2,1,2,1,2,1],[3,3])
    integer :: col_loc(3)
    col_loc = findloc(m, 2, dim=1)
    print *, col_loc(1)
end program test
