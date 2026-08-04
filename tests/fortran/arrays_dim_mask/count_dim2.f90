! vybe-test: fortran/arrays_dim_mask/count_dim2
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    logical :: m(2,3) = reshape([.true.,.false.,.true.,.true.,.false.,.true.],[2,3])
    integer :: row_count(2)
    row_count = count(m, dim=2)
    print *, row_count(1)
end program test
