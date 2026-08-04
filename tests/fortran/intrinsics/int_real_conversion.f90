! vybe-test: fortran/intrinsics/int_real_conversion
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    integer :: n
    real :: x
    n = int(3.7)
    x = real(42)
    print *, n
    print *, x
end program test
