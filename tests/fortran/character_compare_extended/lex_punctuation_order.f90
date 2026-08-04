! vybe-test: fortran/character_compare_extended/lex_punctuation_order
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('.', ',')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", llt('.', ','), "]"
    stop 1
end if
end program t
