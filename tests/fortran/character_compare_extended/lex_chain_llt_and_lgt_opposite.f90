! vybe-test: fortran/character_compare_extended/lex_chain_llt_and_lgt_opposite
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('a','b') .and. lgt('b','a')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('a','b') .and. lgt('b','a'), "]"
    stop 1
end if
end program t
