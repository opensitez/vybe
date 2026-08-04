! vybe-test: fortran/character/char_adjustr_14
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='a    '
print *, adjustr(s)
end program p
