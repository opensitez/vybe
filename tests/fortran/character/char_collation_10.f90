! vybe-test: fortran/character/char_collation_10
! origin: languages/fortran/tests/fortran/test_character.rs
program p
logical :: l
l = 'a' <= 'b'
print *, l
end program p
