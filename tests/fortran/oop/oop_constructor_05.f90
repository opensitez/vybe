! vybe-test: fortran/oop/oop_constructor_05
! origin: languages/fortran/tests/fortran/test_oop.rs
type::t
integer::x
end type t
program p
type(t)::v
v=t(1)
end program p
