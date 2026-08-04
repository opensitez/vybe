! vybe-test: fortran/fortran_catalog_calibrate/bessel_j0
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
print *, nint(bessel_j0(0.0)*100)
end program t
