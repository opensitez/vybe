! vybe-test: fortran/character_operations/chop_01
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=5) :: s='hello'
print *, s(1:2)
end program p
