! vybe-test: fortran/intrinsics_extended/exp_zero
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = exp(0.0)
print *, x
end program t
