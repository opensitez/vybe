! vybe-test: fortran/c_interoperability/c_size_t_16
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
integer(c_size_t) :: x
print *, x
end program p
