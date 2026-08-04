! vybe-test: fortran/oop/oop_inherit_chain_03
! origin: languages/fortran/tests/fortran/test_oop.rs
type::a
integer::x
end type a
type,extends(a)::b
integer::y
end type b
type,extends(b)::c
integer::z
end type c
