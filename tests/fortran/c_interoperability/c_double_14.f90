! vybe-test: fortran/c_interoperability/c_double_14
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
real(c_double) :: x
print *, x
end program p
