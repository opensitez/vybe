! vybe-test: fortran/character_operations/chop_17
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=4) :: a(2)
a = ['ab  ','cd  ']
print *, a
end program p
