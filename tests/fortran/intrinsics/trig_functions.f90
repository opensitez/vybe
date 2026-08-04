! vybe-test: fortran/intrinsics/trig_functions
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    real :: x
    x = sin(1.0)
    x = cos(1.0)
    x = tan(1.0)
    print *, x
end program test
