! vybe-test: fortran/character_assignment/character_assignment_10
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=5) :: s
s = repeat('x',3)
print *, s
end program p
