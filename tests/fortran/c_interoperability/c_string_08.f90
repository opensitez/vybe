! vybe-test: fortran/c_interoperability/c_string_08
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
character(kind=c_char,len=4) :: s
print *, s
end program p
