! vybe-test: fortran/c_interoperability/c_bind_type_07
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
use iso_c_binding
type, bind(c) :: t
 integer(c_int) :: x
end type t
