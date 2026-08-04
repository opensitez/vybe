! vybe-test: fortran/character_operations/chop_16
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=6) :: s
s = 'ab'//'cd'//'ef'
print *, s
end program p
