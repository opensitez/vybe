! vybe-test: fortran/complex/complex_not_equal
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: a = (1.0, 2.0), b = (1.0, 3.0)
    print *, a /= b
end program test
