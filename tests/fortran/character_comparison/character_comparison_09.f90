! vybe-test: fortran/character_comparison/character_comparison_09
! origin: languages/fortran/tests/fortran/test_character_comparison.rs
program p
logical :: l
l='A'/='a'
print *, l
end program p
