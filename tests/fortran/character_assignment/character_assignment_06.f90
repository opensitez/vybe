! vybe-test: fortran/character_assignment/character_assignment_06
! origin: languages/fortran/tests/fortran/test_character_assignment.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
s='abc'
print *, s
end program p
