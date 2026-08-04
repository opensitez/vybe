! vybe-test: fortran/character_collation/character_collation_03
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='A'<'a'
print *, l
end program p
