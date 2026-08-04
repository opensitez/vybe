! vybe-test: fortran/character/char_array_24
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=3) :: a(2)
a = ['abc','def']
print *, a
end program p
