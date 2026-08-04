! vybe-test: fortran/program_units/program_abstract_interface_09
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
abstract interface
subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
