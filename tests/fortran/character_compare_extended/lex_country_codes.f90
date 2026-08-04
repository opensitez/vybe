! vybe-test: fortran/character_compare_extended/lex_country_codes
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('US', 'USA')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('US', 'USA'), "]"
    stop 1
end if
end program t
