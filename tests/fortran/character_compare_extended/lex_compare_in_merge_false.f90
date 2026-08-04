! vybe-test: fortran/character_compare_extended/lex_compare_in_merge_false
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((merge(1, 0, lgt('a', 'b'))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, lgt('a', 'b')), "]"
    stop 1
end if
end program t
