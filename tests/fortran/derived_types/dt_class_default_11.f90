! vybe-test: fortran/derived_types/dt_class_default_11
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
type::t
integer::x
end type t
class(t),allocatable::o
allocate(o)
end program p
