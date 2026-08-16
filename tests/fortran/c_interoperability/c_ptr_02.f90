! vybe-test: fortran/c_interoperability/c_ptr_02
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program driver
use iso_c_binding
type(c_ptr) :: p
end program driver