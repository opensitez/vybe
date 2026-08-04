! vybe-test: fortran/character/char_concat_11
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=4) :: s
s='ab'//'cd'
print *, s
end program p
