! vybe-test: fortran/character_compare_extended/lex_equal_neither_lt_nor_gt
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('eq','eq')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", llt('eq','eq'), "]"
    stop 1
end if
if ((lgt('eq','eq')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", lgt('eq','eq'), "]"
    stop 1
end if
end program t
