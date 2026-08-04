! vybe-test: fortran/character_compare_extended/lex_hex_letter_case
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('a', 'A')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", llt('a', 'A'), "]"
    stop 1
end if
end program t
