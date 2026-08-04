! vybe-test: fortran/math_f2008/log_gamma_vs_log_gamma
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 10.0
    print *, log_gamma(x)
    print *, log(gamma(x))
end program test
