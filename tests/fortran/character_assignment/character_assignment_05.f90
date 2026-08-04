! vybe-test: fortran/character_assignment/character_assignment_05
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=6) :: s='abcdef'
s(2:3)='ZZ'
print *, s
end program p
