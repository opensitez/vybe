! vybe-test: fortran/arrays_dim_mask/size_dim1
! origin: languages/fortran/tests/fortran/test_arrays_dim_mask.rs

program test
    integer :: m(3,4,5)
    print *, size(m, 1)
    print *, size(m, 2)
    print *, size(m, 3)
end program test
