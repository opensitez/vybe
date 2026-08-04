! vybe-test: fortran/program_units/program_main_with_associate_33
! origin: languages/fortran/tests/fortran/test_program_units.rs
program p
integer :: x=1
associate(y=>x)
 print *, y
end associate
end program p
