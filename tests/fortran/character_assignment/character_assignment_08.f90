! vybe-test: fortran/character_assignment/character_assignment_08
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=8) :: s
s='ab'//'cd'
print *, s
end program p
