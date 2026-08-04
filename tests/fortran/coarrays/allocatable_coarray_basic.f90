! vybe-test: fortran/coarrays/allocatable_coarray_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer, allocatable :: x[:]
    allocate(x[*])
    x = this_image()
    print *, x
    deallocate(x)
end program test
