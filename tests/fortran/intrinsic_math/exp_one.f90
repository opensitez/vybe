! vybe-test: fortran/intrinsic_math/exp_one
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
real :: x
x = exp(1.0)
print *, x
end program t
