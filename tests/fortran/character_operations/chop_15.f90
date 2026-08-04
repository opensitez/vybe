! vybe-test: fortran/character_operations/chop_15
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
logical :: l
l = 'A' /= 'a'
print *, l
end program p
