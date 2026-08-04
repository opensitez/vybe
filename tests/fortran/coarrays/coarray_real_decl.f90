! vybe-test: fortran/coarrays/coarray_real_decl
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    real :: r[*]
    r = 3.14
    print *, r
end program test
