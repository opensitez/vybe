! vybe-test: fortran/character_collation/character_collation_04
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='0'<'9'
print *, l
end program p
