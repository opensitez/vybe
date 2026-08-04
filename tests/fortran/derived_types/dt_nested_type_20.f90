! vybe-test: fortran/derived_types/dt_nested_type_20
! origin: languages/fortran/tests/fortran/test_derived_types.rs
type::t1
integer::x
end type t1
type::t2
type(t1)::a
end type t2
