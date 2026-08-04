! vybe-test: fortran/c_interoperability/c_struct_10
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
use iso_c_binding
type, bind(c) :: point
 integer(c_int) :: x
 integer(c_int) :: y
end type point
