! vybe-test: fortran/c_interoperability/c_int_13
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
integer(c_int) :: x
print *, x
end program p
