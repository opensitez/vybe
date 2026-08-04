! vybe-test: fortran/coarrays/coarray_array_decl
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: a(10)[*]
    a = 0
    a(1) = 1
    print *, a(1)
end program test
