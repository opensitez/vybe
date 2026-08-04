! vybe-test: fortran/complex/complex_sub
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: a = (5.0, 6.0), b = (2.0, 3.0), c
    c = a - b
    print *, c
end program test
