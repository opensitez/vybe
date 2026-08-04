! vybe-test: fortran/complex/complex_negate
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: z = (3.0, 4.0)
    complex :: n
    n = -z
    print *, n
end program test
