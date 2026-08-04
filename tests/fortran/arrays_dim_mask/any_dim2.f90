! vybe-test: fortran/arrays_dim_mask/any_dim2
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    logical :: m(3,2) = reshape([.false.,.true.,.false.,.false.,.false.,.false.],[3,2])
    logical :: row_any(3)
    row_any = any(m, dim=2)
    print *, row_any(1)
end program test
