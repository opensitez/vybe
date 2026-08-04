! vybe-test: fortran/character_compare_extended/lex_reverse_transitivity_on_gt
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('c','b') .and. lgt('b','a') .and. lgt('c','a')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('c','b') .and. lgt('b','a') .and. lgt('c','a'), "]"
    stop 1
end if
end program t
