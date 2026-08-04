! vybe-test: fortran/arrays_dim_mask/any_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    logical :: m(2,3) = reshape([.false.,.false.,.true.,.false.,.false.,.false.],[2,3])
    logical :: col_any(3)
    col_any = any(m, dim=1)
    print *, col_any(1)
    print *, col_any(2)
end program test
