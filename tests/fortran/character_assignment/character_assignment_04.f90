! vybe-test: fortran/character_assignment/character_assignment_04
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=5) :: s='hello'
s(1:1)='H'
print *, s
end program p
