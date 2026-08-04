! vybe-test: fortran/c_interoperability/c_array_09
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
integer(c_int) :: a(3)
print *, a
end program p
