! vybe-test: fortran/arrays_dim_mask/all_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    logical :: m(2,3) = reshape([.true.,.true.,.true.,.false.,.true.,.true.],[2,3])
    logical :: col_all(3)
    col_all = all(m, dim=1)
    print *, col_all(1)
    print *, col_all(2)
end program test
