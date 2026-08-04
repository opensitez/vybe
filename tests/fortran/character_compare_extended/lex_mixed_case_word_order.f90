! vybe-test: fortran/character_compare_extended/lex_mixed_case_word_order
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('Fortran', 'fortran')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('Fortran', 'fortran'), "]"
    stop 1
end if
end program t
