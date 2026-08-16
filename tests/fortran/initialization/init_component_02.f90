! vybe-test: fortran/initialization/init_component_02
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
type::t
integer::x=1
end type t
type(t)::v
print *,v%x
end program p
