! vybe-test: fortran/character_compare_extended/lex_underscore_vs_letter
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('a_b', 'ab')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('a_b', 'ab'), "]"
    stop 1
end if
end program t
