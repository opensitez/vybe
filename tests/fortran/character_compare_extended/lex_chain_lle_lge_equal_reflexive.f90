! vybe-test: fortran/character_compare_extended/lex_chain_lle_lge_equal_reflexive
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lle('x','x') .and. lge('x','x')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle('x','x') .and. lge('x','x'), "]"
    stop 1
end if
end program t
