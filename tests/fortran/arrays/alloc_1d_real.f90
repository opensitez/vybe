! vybe-test: fortran/arrays/alloc_1d_real
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    real, allocatable :: v(:)
    allocate(v(3))
    v(1) = 3.14
    print *, v(1)
    deallocate(v)
end program test
