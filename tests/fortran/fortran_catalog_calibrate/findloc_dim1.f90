! vybe-test: fortran/fortran_catalog_calibrate/findloc_dim1
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
print *, findloc(m, 5, dim=1)
end program t
