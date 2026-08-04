! vybe-test: fortran/character_assignment/character_assignment_01
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=4) :: s
s='ab'
print *, s
end program p
