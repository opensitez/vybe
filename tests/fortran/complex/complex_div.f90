! vybe-test: fortran/complex/complex_div
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: a = (1.0, 0.0), b = (2.0, 0.0), c
    c = a / b
    print *, c
end program test
