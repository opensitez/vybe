! vybe-test: fortran/program_units/program_generic_interface_10
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
interface g
module procedure s1
end interface
contains
subroutine s1()
print *, 1
end subroutine s1
end module m
