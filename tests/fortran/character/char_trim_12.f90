! vybe-test: fortran/character/char_trim_12
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='a   '
print *, trim(s)
end program p
