! vybe-test: fortran/character_collation/character_collation_05
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='abc'<'abd'
print *, l
end program p
