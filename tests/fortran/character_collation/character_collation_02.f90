! vybe-test: fortran/character_collation/character_collation_02
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='A'<'B'
print *, l
end program p
