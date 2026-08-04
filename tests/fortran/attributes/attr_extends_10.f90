! vybe-test: fortran/attributes/attr_extends_10
! origin: languages/fortran/tests/fortran/test_attributes.rs
type :: b
integer::x
end type b
type, extends(b) :: c
integer::y
end type c
