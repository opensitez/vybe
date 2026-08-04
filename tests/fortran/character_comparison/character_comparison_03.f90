! vybe-test: fortran/character_comparison/character_comparison_03
! origin: languages/fortran/tests/fortran/test_character_comparison.rs
program p
logical :: l
l='a'<'b'
print *, l
end program p
