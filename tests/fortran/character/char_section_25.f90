! vybe-test: fortran/character/char_section_25
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='hello'
print *, s(:3)
end program p
