! vybe-test: fortran/character/char_newline_23
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=1) :: c
c = achar(10)
print *, ichar(c)
end program p
