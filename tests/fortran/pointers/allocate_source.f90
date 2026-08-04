! vybe-test: fortran/pointers/allocate_source
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, allocatable :: a(:), b(:)
    a = [1, 2, 3, 4, 5]
    allocate(b, source=a)
    print *, b(1)
    deallocate(a, b)
end program test
