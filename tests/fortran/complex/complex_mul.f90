! vybe-test: fortran/complex/complex_mul
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: a = (1.0, 2.0), b = (3.0, 4.0), c
    c = a * b
    print *, c
end program test
