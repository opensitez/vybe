! vybe-test: fortran/fortran_catalog_calibrate/eoshift
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer :: a(4) = [1,2,3,4]
print *, eoshift(a, 1)
end program t
