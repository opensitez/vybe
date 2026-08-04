! vybe-test: fortran/fortran_catalog_calibrate/aimag
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
complex :: z = (4.0, -3.0)
print *, nint(aimag(z))
end program t
