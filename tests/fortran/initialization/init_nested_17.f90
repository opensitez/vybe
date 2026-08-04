! vybe-test: fortran/initialization/init_nested_17
! origin: languages/fortran/tests/fortran/test_initialization.rs
type::u
integer::y=2
end type u
type::t
type(u)::u1
end type t
program p
type(t)::v
end program p
