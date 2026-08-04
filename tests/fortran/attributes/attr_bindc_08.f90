! vybe-test: fortran/attributes/attr_bindc_08
! origin: languages/fortran/tests/fortran/test_attributes.rs
type, bind(c) :: t
integer :: x
end type t
