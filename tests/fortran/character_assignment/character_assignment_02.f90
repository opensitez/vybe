! vybe-test: fortran/character_assignment/character_assignment_02
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=4) :: s='abcd'
s='xy'
print *, s
end program p
