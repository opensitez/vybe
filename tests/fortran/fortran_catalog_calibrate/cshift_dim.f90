! vybe-test: fortran/fortran_catalog_calibrate/cshift_dim
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
print *, cshift(m, 1, dim=2)
end program t
