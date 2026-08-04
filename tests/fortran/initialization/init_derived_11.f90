! vybe-test: fortran/initialization/init_derived_11
! origin: languages/fortran/tests/fortran/test_initialization.rs
type::t
integer::x=1
end type t
program p
type(t)::v=t(2)
print *,v%x
end program p
