! vybe-test: fortran/initialization/init_component_array_21
! origin: languages/fortran/tests/fortran/test_initialization.rs
type::t
integer :: a(3) = [1,2,3]
end type t
program p
type(t)::v
print *, v%a(2)
end program p
