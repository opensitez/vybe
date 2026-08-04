! vybe-test: fortran/coarrays/coarray_2d_array_decl
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    real :: m(4,4)[*]
    m = 0.0
    m(1,1) = 1.0
    print *, m(1,1)
end program test
