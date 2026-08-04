! vybe-test: fortran/character/char_assign_05
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=4) :: s
s='ab'
print *, s
end program p
