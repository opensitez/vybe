! vybe-test: fortran/character_operations/chop_20
! origin: languages/fortran/tests/fortran/test_character_operations.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
s='abc'
print *, s
end program p
