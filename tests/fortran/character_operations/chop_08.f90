! vybe-test: fortran/character_operations/chop_08
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
print *, scan('abc123','0123456789')
end program p
