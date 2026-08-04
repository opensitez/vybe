! vybe-test: fortran/program_units/program_main_with_module_31
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: x=1
end module m
program p
use m
print *, x
end program p
