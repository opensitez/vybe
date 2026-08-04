! vybe-test: fortran/character_compare_extended/lex_year_strings
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('1999', '2000')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('1999', '2000'), "]"
    stop 1
end if
end program t
