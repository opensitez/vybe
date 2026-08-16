! vybe-test: fortran/derived_types/dt_comp_init_18
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x=1
real::y=2.0
end type t
type(t)::v
print *,v%x
end program p
