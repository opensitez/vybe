! vybe-test: fortran/pointers/allocatable_scalar
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, allocatable :: x
    allocate(x)
    x = 42
    print *, x
    deallocate(x)
end program test
