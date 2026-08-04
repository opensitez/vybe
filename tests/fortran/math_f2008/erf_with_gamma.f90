! vybe-test: fortran/math_f2008/erf_with_gamma
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 0.5
    print *, erf(x)
    print *, gamma(x + 0.5) / gamma(0.5)
end program test
