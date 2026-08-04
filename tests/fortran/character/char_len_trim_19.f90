! vybe-test: fortran/character/char_len_trim_19
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=6) :: s='ab    '
print *, len_trim(s)
end program p
