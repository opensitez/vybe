! vybe-test: fortran/initialization/init_structure_ctor_16
! origin: languages/fortran/tests/fortran/test_initialization.rs
type::t
integer::x
end type t
program p
type(t)::v
v=t(1)
end program p
