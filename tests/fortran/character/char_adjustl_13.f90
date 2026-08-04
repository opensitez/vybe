! vybe-test: fortran/character/char_adjustl_13
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='  a  '
print *, adjustl(s)
end program p
