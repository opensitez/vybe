! vybe-test: fortran/complex/real_aimag_both
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: z = (3.0, 4.0)
    real :: r, i
    r = real(z)
    i = aimag(z)
    print *, r
    print *, i
end program test
