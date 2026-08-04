! vybe-test: fortran/character_collation/character_collation_06
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='abc'<='abc'
print *, l
end program p
