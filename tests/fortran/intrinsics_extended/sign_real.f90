! vybe-test: fortran/intrinsics_extended/sign_real
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = sign(3.14, -1.0)
print *, x
end program t
