! vybe-test: fortran/character_compare_extended/lex_eqv_reflexive_lge
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('test','test') .eqv. lle('test','test')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('test','test') .eqv. lle('test','test'), "]"
    stop 1
end if
end program t
