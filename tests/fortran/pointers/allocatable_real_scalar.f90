! vybe-test: fortran/pointers/allocatable_real_scalar
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    real, allocatable :: r
    allocate(r)
    r = 3.14
    print *, r
    deallocate(r)
end program test
