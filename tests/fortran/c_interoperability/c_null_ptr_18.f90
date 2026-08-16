! vybe-test: fortran/c_interoperability/c_null_ptr_18
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program driver
use iso_c_binding
type(c_ptr) :: p
p = c_null_ptr
end program driver