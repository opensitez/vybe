! vybe-test: fortran/intrinsics_extended/atan2_basic
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = atan2(1.0, 1.0)
print *, x
end program t
