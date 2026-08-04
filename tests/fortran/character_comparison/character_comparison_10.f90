! vybe-test: fortran/character_comparison/character_comparison_10
! origin: languages/fortran/tests/fortran/test_character_comparison.rs
program p
character(len=3) :: a='abc', b='abd'
print *, a < b
end program p
