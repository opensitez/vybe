! vybe-test: fortran/c_interoperability/c_loc_04
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
integer, target :: x
print *, c_associated(c_loc(x))
end program p
