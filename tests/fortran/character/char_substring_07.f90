! vybe-test: fortran/character/char_substring_07
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='hello'
print *, s(2:4)
end program p
