! vybe-test: fortran/program_units/program_contains_02
! origin: languages/fortran/tests/fortran/test_program_units.rs
program p
contains
subroutine s()
print *, 1
end subroutine s
end program p
