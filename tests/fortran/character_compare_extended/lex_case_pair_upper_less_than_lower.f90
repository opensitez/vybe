! vybe-test: fortran/character_compare_extended/lex_case_pair_upper_less_than_lower
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('CAT', 'cat')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('CAT', 'cat'), "]"
    stop 1
end if
end program t
