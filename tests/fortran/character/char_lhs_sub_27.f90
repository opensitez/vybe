! vybe-test: fortran/character/char_lhs_sub_27
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=5) :: s='hello'
s(1:1)='H'
print *, s
end program p
