! vybe-test: fortran/c_interoperability/c_associated_20
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program driver
use iso_c_binding
type(c_ptr) :: p
print *, c_associated(p)
end program driver