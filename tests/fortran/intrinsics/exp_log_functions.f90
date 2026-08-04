! vybe-test: fortran/intrinsics/exp_log_functions
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    real :: x
    x = exp(1.0)
    x = log(2.718)
    print *, x
end program test
