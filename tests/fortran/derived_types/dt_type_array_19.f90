! vybe-test: fortran/derived_types/dt_type_array_19
! origin: languages/fortran/tests/fortran/test_derived_types.rs
type::t
integer::x
end type t
program p
type(t)::a(2)
print *,1
end program p
