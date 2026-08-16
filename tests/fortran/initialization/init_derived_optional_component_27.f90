! vybe-test: fortran/initialization/init_derived_optional_component_27
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
type::inner
integer :: x = 10
end type inner
type::outer
type(inner) :: c
integer :: y = 4
end type outer
type(outer)::v
print *, v%c%x
print *, v%y
end program p
