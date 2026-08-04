! vybe-test: fortran/character_comparison/character_comparison_08
! origin: languages/fortran/tests/fortran/test_character_comparison.rs
program p
logical :: l
l='abc'/='abd'
print *, l
end program p
