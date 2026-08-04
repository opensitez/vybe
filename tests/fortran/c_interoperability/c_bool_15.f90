! vybe-test: fortran/c_interoperability/c_bool_15
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
logical(c_bool) :: x
print *, x
end program p
