! vybe-test: fortran/derived_types/dt_constructor_16
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x
end type t
type(t)::v
v=t(1)
print *,v%x
end program p
