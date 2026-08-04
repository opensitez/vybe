! vybe-test: fortran/derived_types/dt_type_ext_01
! origin: languages/fortran/tests/fortran/test_derived_types.rs
type::b
integer::x
end type b
type,extends(b)::c
integer::y
end type c
