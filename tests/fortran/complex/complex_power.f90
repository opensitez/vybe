! vybe-test: fortran/complex/complex_power
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: z = (1.0, 1.0)
    complex :: r
    r = z ** 2
    print *, r
end program test
