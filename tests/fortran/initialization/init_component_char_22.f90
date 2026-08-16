! vybe-test: fortran/initialization/init_component_char_22
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
type::t
character(len=4) :: tag = 'init'
logical :: active = .true.
end type t
type(t)::v
print *, v%tag
print *, v%active
end program p
