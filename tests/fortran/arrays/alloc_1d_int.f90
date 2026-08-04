! vybe-test: fortran/arrays/alloc_1d_int
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer, allocatable :: v(:)
    allocate(v(5))
    v(1) = 42
    print *, v(1)
    deallocate(v)
end program test
