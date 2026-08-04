! vybe-test: fortran/fortran_character_quality/character_quality_compare_equality
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_compare_equality
    character(len=6) :: left
    character(len=6) :: right
    left = 'abc   '
    right = 'abc   '
    if (left == right) print *, 1
    if (left /= right) print *, 0
end program character_quality_compare_equality
