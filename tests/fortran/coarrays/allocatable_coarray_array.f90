! vybe-test: fortran/coarrays/allocatable_coarray_array
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    real, allocatable :: a(:)[:]
    allocate(a(10)[*])
    a = real(this_image())
    print *, a(1)
    deallocate(a)
end program test
