! vybe-test: fortran/character/char_rhs_expr_26
! origin: languages/fortran/tests/fortran/test_character.rs
program p
character(len=6) :: s
s = adjustl('  ab')
print *, s
end program p
