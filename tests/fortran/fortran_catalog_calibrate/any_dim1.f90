! vybe-test: fortran/fortran_catalog_calibrate/any_dim1
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
logical :: m(2,2) = reshape([.false.,.true.,.false.,.false.],[2,2])
logical :: c(2)
c = any(m, dim=1)
print *, c(1)
end program t
