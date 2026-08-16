! vybe-test: fortran/oop/oop_lifetime_10
! origin: languages/fortran/tests/fortran/test_oop.rs
program p
type::t
integer::x
end type t
block
type(t)::v
v%x=1
end block
end program p
