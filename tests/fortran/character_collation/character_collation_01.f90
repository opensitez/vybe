! vybe-test: fortran/character_collation/character_collation_01
! origin: languages/fortran/tests/fortran/test_character_collation.rs
program p
logical :: l
l='a'<'b'
print *, l
end program p
