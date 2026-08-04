! vybe-test: fortran/character_operations/chop_02
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=5) :: s='hello'
s(2:3)='ZZ'
print *, s
end program p
