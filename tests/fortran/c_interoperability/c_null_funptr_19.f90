! vybe-test: fortran/c_interoperability/c_null_funptr_19
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
type(c_funptr) :: fp
fp = c_null_funptr
end program p
