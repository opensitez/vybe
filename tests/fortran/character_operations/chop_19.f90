! vybe-test: fortran/character_operations/chop_19
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=20) :: buf='42'
integer :: x
read(buf,*) x
print *, x
end program p
