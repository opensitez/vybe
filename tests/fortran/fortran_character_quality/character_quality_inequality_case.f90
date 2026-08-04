! vybe-test: fortran/fortran_character_quality/character_quality_inequality_case
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_inequality_case
    character(len=6) :: a
    character(len=6) :: b
    a = 'abc   '
    b = 'abd   '
    if (a < b) print *, 1
    if (a > b) print *, 0
end program character_quality_inequality_case
