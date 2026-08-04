! vybe-test: fortran/character_compare_extended/lex_transitivity_chain
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('a','b') .and. llt('b','c') .and. llt('a','c')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('a','b') .and. llt('b','c') .and. llt('a','c'), "]"
    stop 1
end if
end program t
