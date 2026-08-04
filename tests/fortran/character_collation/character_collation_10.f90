! vybe-test: fortran/character_collation/character_collation_10
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='same'=='same'
print *, l
end program p
