! vybe-test: fortran/complex/complex_mixed_real
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: z = (3.0, 4.0)
    complex :: r
    r = 2.0 * z
    print *, r
end program test
