! vybe-test: fortran/character_collation/character_collation_08
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='b'>'a'
print *, l
end program p
