! vybe-test: fortran/character_assignment/character_assignment_09
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=5) :: s='abcde'
s = adjustl('  xy')
print *, s
end program p
