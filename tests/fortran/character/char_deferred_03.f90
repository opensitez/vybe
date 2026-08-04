! vybe-test: fortran/character/char_deferred_03
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
end program p
