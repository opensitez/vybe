! vybe-test: fortran/derived_types/dt_default_init_17
! origin: languages/fortran/tests/fortran/test_derived_types.rs
type::t
integer::x=1
end type t
program p
type(t)::v
print *,v%x
end program p
