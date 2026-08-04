! vybe-test: fortran/c_interoperability/c_char_17
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
character(kind=c_char) :: x
print *, x
end program p
