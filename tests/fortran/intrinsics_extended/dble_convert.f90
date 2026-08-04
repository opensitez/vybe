! vybe-test: fortran/intrinsics_extended/dble_convert
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
double precision :: d
d = dble(3)
print *, d
end program t
