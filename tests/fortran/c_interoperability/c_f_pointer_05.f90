! vybe-test: fortran/c_interoperability/c_f_pointer_05
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program driver
use iso_c_binding
type(c_ptr) :: p
integer, pointer :: fp
call c_f_pointer(p, fp)
end program driver