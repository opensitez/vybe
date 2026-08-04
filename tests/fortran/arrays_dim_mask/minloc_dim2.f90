! vybe-test: fortran/arrays_dim_mask/minloc_dim2
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
    integer :: row_minloc(3)
    row_minloc = minloc(m, dim=2)
    print *, row_minloc(1)
end program test
