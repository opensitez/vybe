! vybe-test: fortran/arrays_dim_mask/all_dim2
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    logical :: m(3,2) = reshape([.true.,.true.,.true.,.true.,.false.,.true.],[3,2])
    logical :: row_all(3)
    row_all = all(m, dim=2)
    print *, row_all(1)
end program test
