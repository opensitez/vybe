! vybe-test: fortran/character_operations/chop_14
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
logical :: l
l = 'a' < 'b'
print *, l
end program p
