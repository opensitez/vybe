! vybe-test: fortran/program_units/program_assignment_interface_12
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
interface assignment(=)
module procedure assigni
end interface
contains
subroutine assigni(a,b)
integer :: a,b
a = b
end subroutine assigni
end module m
