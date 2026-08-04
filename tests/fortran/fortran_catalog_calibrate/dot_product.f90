! vybe-test: fortran/fortran_catalog_calibrate/dot_product
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer :: a(3) = [1,2,3]
integer :: b(3) = [4,5,6]
print *, dot_product(a,b)
end program t
