! vybe-test: fortran/fortran_catalog_calibrate/lgamma
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
print *, nint(lgamma(2.0)*100)
end program t
