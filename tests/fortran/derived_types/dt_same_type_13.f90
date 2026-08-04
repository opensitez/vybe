! vybe-test: fortran/derived_types/dt_same_type_13
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x
end type t
type(t)::a,b
print *, same_type_as(a,b)
end program p
