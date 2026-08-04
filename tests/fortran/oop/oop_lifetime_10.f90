! vybe-test: fortran/oop/oop_lifetime_10
! origin: languages/fortran/tests/fortran/test_oop.rs
type::t
integer::x
end type t
program p
block
type(t)::v
v%x=1
end block
end program p
