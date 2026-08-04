! vybe-test: fortran/fortran_catalog_calibrate/allocated
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer, allocatable :: a(:)
print *, allocated(a)
allocate(a(2))
print *, allocated(a)
end program t
