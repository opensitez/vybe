! vybe-test: fortran/fortran_catalog_calibrate/count_dim1
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer :: a(2,3) = reshape([1,2,3,4,5,6],[2,3])
print *, count(a > 3, dim=1)
end program t
