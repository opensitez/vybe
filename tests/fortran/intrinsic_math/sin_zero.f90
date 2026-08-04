! vybe-test: fortran/intrinsic_math/sin_zero
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
real :: x
x = sin(0.0)
print *, x
end program t
