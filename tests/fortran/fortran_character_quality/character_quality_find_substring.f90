! vybe-test: fortran/fortran_character_quality/character_quality_find_substring
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_find_substring
    character(len=20) :: source
    source = 'fortran language'
    if ((index(source, 'lang')) /= 9) then
    print *, "FAIL: want [9] got [", index(source, 'lang'), "]"
    stop 1
end if
    if ((index(source, 'xx')) /= 0) then
    print *, "FAIL: want [0] got [", index(source, 'xx'), "]"
    stop 1
end if
end program character_quality_find_substring
