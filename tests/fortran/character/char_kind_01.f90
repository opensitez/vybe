! vybe-test: fortran/character/char_kind_01
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(kind=1,len=4) :: s
print *, s
end program p
