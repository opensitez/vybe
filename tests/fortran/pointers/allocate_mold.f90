! vybe-test: fortran/pointers/allocate_mold
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    integer, allocatable :: a(:), b(:)
    allocate(a(5))
    allocate(b, mold=a)
    print *, size(b)
    deallocate(a, b)
end program test
