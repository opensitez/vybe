! vybe-test: fortran/character_assignment/character_assignment_07
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=4) :: a(2)
a(1)='ab'
a(2)='cd'
print *, a
end program p
