! vybe-test: fortran/fortran2003/allocated_true
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    integer, allocatable :: x(:)
    allocate(x(5))
    print *, allocated(x)
    deallocate(x)
end program test
