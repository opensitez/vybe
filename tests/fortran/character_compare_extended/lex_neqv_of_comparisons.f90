! vybe-test: fortran/character_compare_extended/lex_neqv_of_comparisons
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('a','b') .neqv. lgt('a','b')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('a','b') .neqv. lgt('a','b'), "]"
    stop 1
end if
end program t
