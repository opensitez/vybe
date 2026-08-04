! vybe-test: fortran/fortran_character_quality/character_quality_index_case_sensitive
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_index_case_sensitive
    character(len=20) :: source
    source = 'Fortran Fortran'
    if ((index(source, 'For')) /= 1) then
    print *, "FAIL: want [1] got [", index(source, 'For'), "]"
    stop 1
end if
    if ((index(source, 'for')) /= 0) then
    print *, "FAIL: want [0] got [", index(source, 'for'), "]"
    stop 1
end if
end program character_quality_index_case_sensitive
