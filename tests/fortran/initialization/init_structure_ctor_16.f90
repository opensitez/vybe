! vybe-test: fortran/initialization/init_structure_ctor_16
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
type::t
integer::x
end type t
type(t)::v
v=t(1)
end program p
