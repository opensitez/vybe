! vybe-test: fortran/legacy/intrinsic_decl
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    intrinsic :: sin, cos, sqrt
    real :: x
    x = sin(0.0)
    print *, x
end program test
