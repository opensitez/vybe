! vybe-test: fortran/derived_types/dt_default_init_17
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x=1
end type t
type(t)::v
print *,v%x
end program p
