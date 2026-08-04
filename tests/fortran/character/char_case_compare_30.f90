! vybe-test: fortran/character/char_case_compare_30
! origin: languages/fortran/tests/fortran/test_character.rs
program p
logical :: l
l = 'A' /= 'a'
print *, l
end program p
