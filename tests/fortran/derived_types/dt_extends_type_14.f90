! vybe-test: fortran/derived_types/dt_extends_type_14
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x
end type t
type(t)::a,b
print *, extends_type_of(a,b)
end program p
