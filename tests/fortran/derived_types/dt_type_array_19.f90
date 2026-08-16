! vybe-test: fortran/derived_types/dt_type_array_19
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x
end type t
type(t)::a(2)
print *,1
end program p
