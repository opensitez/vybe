! vybe-test: fortran/c_interoperability/c_funptr_03
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
type(c_funptr) :: fp
end program p
